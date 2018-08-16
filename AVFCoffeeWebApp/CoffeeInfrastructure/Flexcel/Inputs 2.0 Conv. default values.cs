using FlexCel.Core;

namespace CoffeeInfrastructure.Flexcel
{
    public class Inputs_2
    {
        public void Inputs_2_Default(ExcelFile xls)
        {
            xls.NewFile(38, TExcelFileFormat.v2016);    //Create a new Excel file with 38 sheets.

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

            xls.ActiveSheet = 17;    //Set the sheet we are working in.

            //Global Workbook Options
            xls.OptionsAutoCompressPictures = false;
            xls.OptionsMultithreadRecalc = 0;

            //Sheet Options
            xls.SheetName = "Inputs 2.0 Conv. default values";

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

            xls.SetColWidth(2, 2, 22176);    //(85.88 + 0.75) * 256

            xls.SetColWidth(3, 3, 4640);    //(17.38 + 0.75) * 256

            xls.SetColWidth(4, 7, 2784);    //(10.13 + 0.75) * 256

            xls.SetColWidth(8, 8, 2784);    //(10.13 + 0.75) * 256

            TFlxFormat ColFmt;
            ColFmt = xls.GetFormat(xls.GetColFormat(8));
            ColFmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            ColFmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetColFormat(8, 8, xls.AddFormat(ColFmt));

            xls.SetColWidth(9, 9, 4384);    //(16.38 + 0.75) * 256

            xls.SetColWidth(10, 16384, 2784);    //(10.13 + 0.75) * 256
            xls.DefaultRowHeight = 315;

            xls.SetRowHeight(4, 390);    //19.50 * 20
            xls.SetRowHeight(9, 630);    //31.50 * 20
            xls.SetRowHeight(40, 330);    //16.50 * 20
            xls.SetRowHeight(60, 630);    //31.50 * 20
            xls.SetRowHeight(84, 630);    //31.50 * 20
            xls.SetRowHeight(108, 630);    //31.50 * 20
            xls.SetRowHeight(122, 630);    //31.50 * 20
            xls.SetRowHeight(126, 630);    //31.50 * 20
            xls.SetRowHeight(129, 630);    //31.50 * 20
            xls.SetRowHeight(133, 630);    //31.50 * 20
            xls.SetRowHeight(154, 375);    //18.75 * 20
            xls.SetRowHeight(169, 1020);    //51.00 * 20
            xls.SetRowHeight(268, 630);    //31.50 * 20
            xls.SetRowHeight(269, 945);    //47.25 * 20
            xls.SetRowHeight(270, 630);    //31.50 * 20
            xls.SetRowHeight(271, 630);    //31.50 * 20
            xls.SetRowHeight(273, 630);    //31.50 * 20
            xls.SetRowHeight(293, 330);    //16.50 * 20

            //Merged Cells
            xls.MergeCells(4, 2, 4, 3);
            xls.MergeCells(10, 5, 10, 7);
            xls.MergeCells(162, 2, 162, 3);

            //Set the cell values
            TFlxFormat fmt;
            fmt = xls.GetCellVisibleFormatDef(1, 2);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0x33, 0x66, 0xFF);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            fmt.Indent = 1;
            xls.SetCellFormat(1, 2, xls.AddFormat(fmt));
            xls.SetCellValue(1, 2, "Default esta originalmente en: ");
            xls.SetCellValue(1, 3, "Quintales");
            xls.SetCellValue(1, 8, "To user:");
            xls.SetCellValue(1, 9, new TFormula("='Gral Conf. Summary'!$H$15"));
            xls.SetCellValue(2, 3, "Mexican pesos");
            xls.SetCellValue(2, 9, new TFormula("='Gral Conf. Summary'!$H$33"));
            xls.SetCellValue(3, 3, "Hectares");
            xls.SetCellValue(3, 9, new TFormula("='Gral Conf. Summary'!$H$23"));

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
            xls.SetCellValue(9, 2, "Example: If one activity requires 2 people, working 1 day and this activity is performed"
            + " 3 times per year,  then total days = 2*1*3 =6");

            fmt = xls.GetCellVisibleFormatDef(9, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(9, 3, xls.AddFormat(fmt));

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
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(10, 5, xls.AddFormat(fmt));
            xls.SetCellValue(10, 5, "Conversion metrics");

            fmt = xls.GetCellVisibleFormatDef(10, 6);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(10, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(10, 7);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(10, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(10, 8);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(10, 8, xls.AddFormat(fmt));
            xls.SetCellValue(10, 8, "Factor");

            fmt = xls.GetCellVisibleFormatDef(10, 9);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(10, 9, xls.AddFormat(fmt));
            xls.SetCellValue(10, 9, "Input");

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

            fmt = xls.GetCellVisibleFormatDef(11, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(11, 9, xls.AddFormat(fmt));

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
            xls.SetCellValue(12, 3, "days/ha");

            fmt = xls.GetCellVisibleFormatDef(12, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(12, 9, xls.AddFormat(fmt));

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
            xls.SetCellValue(13, 3, 1.71666666666667);
            xls.SetCellValue(13, 5, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(13, 6, 1);
            xls.SetCellValue(13, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(13, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(13, 8, xls.AddFormat(fmt));
            xls.SetCellValue(13, 8, new TFormula("=(    IF(E13<>1,VLOOKUP(E13,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)   )     "
            + " *       IF(F13<>1,VLOOKUP(F13,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) * IF(G13<>1,VLOOKUP(G13,'Gral"
            + " Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(13, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(13, 9, xls.AddFormat(fmt));
            xls.SetCellValue(13, 9, new TFormula("=C13*H13"));

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
            xls.SetCellValue(14, 3, 1.52243333333333);
            xls.SetCellValue(14, 5, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(14, 6, 1);
            xls.SetCellValue(14, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(14, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(14, 8, xls.AddFormat(fmt));
            xls.SetCellValue(14, 8, new TFormula("=(    IF(E14<>1,VLOOKUP(E14,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)   )     "
            + " *       IF(F14<>1,VLOOKUP(F14,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) * IF(G14<>1,VLOOKUP(G14,'Gral"
            + " Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(14, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(14, 9, xls.AddFormat(fmt));
            xls.SetCellValue(14, 9, new TFormula("=C14*H14"));

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
            xls.SetCellValue(15, 3, 4.02777777777778);
            xls.SetCellValue(15, 5, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(15, 6, 1);
            xls.SetCellValue(15, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(15, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(15, 8, xls.AddFormat(fmt));
            xls.SetCellValue(15, 8, new TFormula("=(    IF(E15<>1,VLOOKUP(E15,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)   )     "
            + " *       IF(F15<>1,VLOOKUP(F15,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) * IF(G15<>1,VLOOKUP(G15,'Gral"
            + " Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(15, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(15, 9, xls.AddFormat(fmt));
            xls.SetCellValue(15, 9, new TFormula("=C15*H15"));

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
            xls.SetCellValue(16, 3, 8.82);
            xls.SetCellValue(16, 5, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(16, 6, 1);
            xls.SetCellValue(16, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(16, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(16, 8, xls.AddFormat(fmt));
            xls.SetCellValue(16, 8, new TFormula("=(    IF(E16<>1,VLOOKUP(E16,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)   )     "
            + " *       IF(F16<>1,VLOOKUP(F16,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) * IF(G16<>1,VLOOKUP(G16,'Gral"
            + " Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(16, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(16, 9, xls.AddFormat(fmt));
            xls.SetCellValue(16, 9, new TFormula("=C16*H16"));

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
            xls.SetCellValue(17, 3, 0);
            xls.SetCellValue(17, 5, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(17, 6, 1);
            xls.SetCellValue(17, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(17, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(17, 8, xls.AddFormat(fmt));
            xls.SetCellValue(17, 8, new TFormula("=(    IF(E17<>1,VLOOKUP(E17,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)   )     "
            + " *       IF(F17<>1,VLOOKUP(F17,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) * IF(G17<>1,VLOOKUP(G17,'Gral"
            + " Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(17, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(17, 9, xls.AddFormat(fmt));
            xls.SetCellValue(17, 9, new TFormula("=C17*H17"));

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

            fmt = xls.GetCellVisibleFormatDef(18, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(18, 9, xls.AddFormat(fmt));

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
            xls.SetCellValue(19, 3, 9.61224489795918);
            xls.SetCellValue(19, 5, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(19, 6, 1);
            xls.SetCellValue(19, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(19, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(19, 8, xls.AddFormat(fmt));
            xls.SetCellValue(19, 8, new TFormula("=(    IF(E19<>1,VLOOKUP(E19,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)   )     "
            + " *       IF(F19<>1,VLOOKUP(F19,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) * IF(G19<>1,VLOOKUP(G19,'Gral"
            + " Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(19, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(19, 9, xls.AddFormat(fmt));
            xls.SetCellValue(19, 9, new TFormula("=C19*H19"));

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
            xls.SetCellValue(20, 3, 8.92);
            xls.SetCellValue(20, 5, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(20, 6, 1);
            xls.SetCellValue(20, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(20, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(20, 8, xls.AddFormat(fmt));
            xls.SetCellValue(20, 8, new TFormula("=(    IF(E20<>1,VLOOKUP(E20,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)   )     "
            + " *       IF(F20<>1,VLOOKUP(F20,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) * IF(G20<>1,VLOOKUP(G20,'Gral"
            + " Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(20, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(20, 9, xls.AddFormat(fmt));
            xls.SetCellValue(20, 9, new TFormula("=C20*H20"));

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
            xls.SetCellValue(21, 3, 16.9833333333333);
            xls.SetCellValue(21, 5, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(21, 6, 1);
            xls.SetCellValue(21, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(21, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(21, 8, xls.AddFormat(fmt));
            xls.SetCellValue(21, 8, new TFormula("=(    IF(E21<>1,VLOOKUP(E21,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)   )     "
            + " *       IF(F21<>1,VLOOKUP(F21,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) * IF(G21<>1,VLOOKUP(G21,'Gral"
            + " Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(21, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(21, 9, xls.AddFormat(fmt));
            xls.SetCellValue(21, 9, new TFormula("=C21*H21"));

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
            xls.SetCellValue(22, 3, 6.3366);
            xls.SetCellValue(22, 5, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(22, 6, 1);
            xls.SetCellValue(22, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(22, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(22, 8, xls.AddFormat(fmt));
            xls.SetCellValue(22, 8, new TFormula("=(    IF(E22<>1,VLOOKUP(E22,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)   )     "
            + " *       IF(F22<>1,VLOOKUP(F22,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) * IF(G22<>1,VLOOKUP(G22,'Gral"
            + " Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(22, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(22, 9, xls.AddFormat(fmt));
            xls.SetCellValue(22, 9, new TFormula("=C22*H22"));

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
            xls.SetCellValue(23, 3, 14.78);
            xls.SetCellValue(23, 5, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(23, 6, 1);
            xls.SetCellValue(23, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(23, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(23, 8, xls.AddFormat(fmt));
            xls.SetCellValue(23, 8, new TFormula("=(    IF(E23<>1,VLOOKUP(E23,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)   )     "
            + " *       IF(F23<>1,VLOOKUP(F23,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) * IF(G23<>1,VLOOKUP(G23,'Gral"
            + " Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(23, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(23, 9, xls.AddFormat(fmt));
            xls.SetCellValue(23, 9, new TFormula("=C23*H23"));

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
            xls.SetCellValue(24, 3, 5.45);
            xls.SetCellValue(24, 5, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(24, 6, 1);
            xls.SetCellValue(24, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(24, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(24, 8, xls.AddFormat(fmt));
            xls.SetCellValue(24, 8, new TFormula("=(    IF(E24<>1,VLOOKUP(E24,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)   )     "
            + " *       IF(F24<>1,VLOOKUP(F24,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) * IF(G24<>1,VLOOKUP(G24,'Gral"
            + " Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(24, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(24, 9, xls.AddFormat(fmt));
            xls.SetCellValue(24, 9, new TFormula("=C24*H24"));

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
            xls.SetCellValue(25, 3, 24.5273333333333);
            xls.SetCellValue(25, 5, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(25, 6, 1);
            xls.SetCellValue(25, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(25, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(25, 8, xls.AddFormat(fmt));
            xls.SetCellValue(25, 8, new TFormula("=(    IF(E25<>1,VLOOKUP(E25,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)   )     "
            + " *       IF(F25<>1,VLOOKUP(F25,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) * IF(G25<>1,VLOOKUP(G25,'Gral"
            + " Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(25, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(25, 9, xls.AddFormat(fmt));
            xls.SetCellValue(25, 9, new TFormula("=C25*H25"));

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
            xls.SetCellValue(26, 3, 2.41153333333333);
            xls.SetCellValue(26, 5, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(26, 6, 1);
            xls.SetCellValue(26, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(26, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(26, 8, xls.AddFormat(fmt));
            xls.SetCellValue(26, 8, new TFormula("=(    IF(E26<>1,VLOOKUP(E26,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)   )     "
            + " *       IF(F26<>1,VLOOKUP(F26,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) * IF(G26<>1,VLOOKUP(G26,'Gral"
            + " Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(26, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(26, 9, xls.AddFormat(fmt));
            xls.SetCellValue(26, 9, new TFormula("=C26*H26"));

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
            xls.SetCellValue(27, 3, 1.44444444444444);
            xls.SetCellValue(27, 5, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(27, 6, 1);
            xls.SetCellValue(27, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(27, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(27, 8, xls.AddFormat(fmt));
            xls.SetCellValue(27, 8, new TFormula("=(    IF(E27<>1,VLOOKUP(E27,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)   )     "
            + " *       IF(F27<>1,VLOOKUP(F27,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) * IF(G27<>1,VLOOKUP(G27,'Gral"
            + " Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(27, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(27, 9, xls.AddFormat(fmt));
            xls.SetCellValue(27, 9, new TFormula("=C27*H27"));

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
            xls.SetCellValue(28, 3, 0.3);
            xls.SetCellValue(28, 5, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(28, 6, 1);
            xls.SetCellValue(28, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(28, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(28, 8, xls.AddFormat(fmt));
            xls.SetCellValue(28, 8, new TFormula("=(    IF(E28<>1,VLOOKUP(E28,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)   )     "
            + " *       IF(F28<>1,VLOOKUP(F28,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) * IF(G28<>1,VLOOKUP(G28,'Gral"
            + " Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(28, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(28, 9, xls.AddFormat(fmt));
            xls.SetCellValue(28, 9, new TFormula("=C28*H28"));

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

            fmt = xls.GetCellVisibleFormatDef(29, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(29, 9, xls.AddFormat(fmt));

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
            xls.SetCellValue(30, 3, 18.78);
            xls.SetCellValue(30, 5, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(30, 6, 1);
            xls.SetCellValue(30, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(30, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(30, 8, xls.AddFormat(fmt));
            xls.SetCellValue(30, 8, new TFormula("=(    IF(E30<>1,VLOOKUP(E30,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)   )     "
            + " *       IF(F30<>1,VLOOKUP(F30,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) * IF(G30<>1,VLOOKUP(G30,'Gral"
            + " Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(30, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(30, 9, xls.AddFormat(fmt));
            xls.SetCellValue(30, 9, new TFormula("=C30*H30"));

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
            xls.SetCellValue(31, 3, 13.48);
            xls.SetCellValue(31, 5, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(31, 6, 1);
            xls.SetCellValue(31, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(31, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(31, 8, xls.AddFormat(fmt));
            xls.SetCellValue(31, 8, new TFormula("=(    IF(E31<>1,VLOOKUP(E31,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)   )     "
            + " *       IF(F31<>1,VLOOKUP(F31,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) * IF(G31<>1,VLOOKUP(G31,'Gral"
            + " Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(31, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(31, 9, xls.AddFormat(fmt));
            xls.SetCellValue(31, 9, new TFormula("=C31*H31"));

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
            xls.SetCellValue(32, 3, 3.5);
            xls.SetCellValue(32, 5, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(32, 6, 1);
            xls.SetCellValue(32, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(32, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(32, 8, xls.AddFormat(fmt));
            xls.SetCellValue(32, 8, new TFormula("=(    IF(E32<>1,VLOOKUP(E32,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)   )     "
            + " *       IF(F32<>1,VLOOKUP(F32,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) * IF(G32<>1,VLOOKUP(G32,'Gral"
            + " Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(32, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(32, 9, xls.AddFormat(fmt));
            xls.SetCellValue(32, 9, new TFormula("=C32*H32"));

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
            xls.SetCellValue(33, 3, 6.12);
            xls.SetCellValue(33, 5, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(33, 6, 1);
            xls.SetCellValue(33, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(33, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(33, 8, xls.AddFormat(fmt));
            xls.SetCellValue(33, 8, new TFormula("=(    IF(E33<>1,VLOOKUP(E33,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)   )     "
            + " *       IF(F33<>1,VLOOKUP(F33,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) * IF(G33<>1,VLOOKUP(G33,'Gral"
            + " Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(33, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(33, 9, xls.AddFormat(fmt));
            xls.SetCellValue(33, 9, new TFormula("=C33*H33"));

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
            xls.SetCellValue(34, 3, 10.78);
            xls.SetCellValue(34, 5, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(34, 6, 1);
            xls.SetCellValue(34, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(34, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(34, 8, xls.AddFormat(fmt));
            xls.SetCellValue(34, 8, new TFormula("=(    IF(E34<>1,VLOOKUP(E34,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)   )     "
            + " *       IF(F34<>1,VLOOKUP(F34,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) * IF(G34<>1,VLOOKUP(G34,'Gral"
            + " Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(34, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(34, 9, xls.AddFormat(fmt));
            xls.SetCellValue(34, 9, new TFormula("=C34*H34"));

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
            xls.SetCellValue(35, 3, 27.38);
            xls.SetCellValue(35, 5, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(35, 6, 1);
            xls.SetCellValue(35, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(35, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(35, 8, xls.AddFormat(fmt));
            xls.SetCellValue(35, 8, new TFormula("=(    IF(E35<>1,VLOOKUP(E35,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)   )     "
            + " *       IF(F35<>1,VLOOKUP(F35,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) * IF(G35<>1,VLOOKUP(G35,'Gral"
            + " Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(35, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(35, 9, xls.AddFormat(fmt));
            xls.SetCellValue(35, 9, new TFormula("=C35*H35"));

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
            xls.SetCellValue(36, 3, 12.9013333333333);
            xls.SetCellValue(36, 5, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(36, 6, 1);
            xls.SetCellValue(36, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(36, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(36, 8, xls.AddFormat(fmt));
            xls.SetCellValue(36, 8, new TFormula("=(    IF(E36<>1,VLOOKUP(E36,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)   )     "
            + " *       IF(F36<>1,VLOOKUP(F36,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) * IF(G36<>1,VLOOKUP(G36,'Gral"
            + " Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(36, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(36, 9, xls.AddFormat(fmt));
            xls.SetCellValue(36, 9, new TFormula("=C36*H36"));

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
            xls.SetCellValue(37, 3, 23.34);
            xls.SetCellValue(37, 5, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(37, 6, 1);
            xls.SetCellValue(37, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(37, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(37, 8, xls.AddFormat(fmt));
            xls.SetCellValue(37, 8, new TFormula("=(    IF(E37<>1,VLOOKUP(E37,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)   )     "
            + " *       IF(F37<>1,VLOOKUP(F37,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) * IF(G37<>1,VLOOKUP(G37,'Gral"
            + " Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(37, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(37, 9, xls.AddFormat(fmt));
            xls.SetCellValue(37, 9, new TFormula("=C37*H37"));

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
            xls.SetCellValue(38, 3, 13.32);
            xls.SetCellValue(38, 5, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(38, 6, 1);
            xls.SetCellValue(38, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(38, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(38, 8, xls.AddFormat(fmt));
            xls.SetCellValue(38, 8, new TFormula("=(    IF(E38<>1,VLOOKUP(E38,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)   )     "
            + " *       IF(F38<>1,VLOOKUP(F38,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) * IF(G38<>1,VLOOKUP(G38,'Gral"
            + " Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(38, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(38, 9, xls.AddFormat(fmt));
            xls.SetCellValue(38, 9, new TFormula("=C38*H38"));

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
            xls.SetCellValue(39, 3, 4.66);
            xls.SetCellValue(39, 5, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(39, 6, 1);
            xls.SetCellValue(39, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(39, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(39, 8, xls.AddFormat(fmt));
            xls.SetCellValue(39, 8, new TFormula("=(    IF(E39<>1,VLOOKUP(E39,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)   )     "
            + " *       IF(F39<>1,VLOOKUP(F39,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) * IF(G39<>1,VLOOKUP(G39,'Gral"
            + " Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(39, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(39, 9, xls.AddFormat(fmt));
            xls.SetCellValue(39, 9, new TFormula("=C39*H39"));

            fmt = xls.GetCellVisibleFormatDef(40, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(40, 2, xls.AddFormat(fmt));
            xls.SetCellValue(40, 2, "Others:");

            fmt = xls.GetCellVisibleFormatDef(40, 3);
            fmt.Font.Family = 0;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.Format = "0.00";
            xls.SetCellFormat(40, 3, xls.AddFormat(fmt));
            xls.SetCellValue(40, 3, 1.2998);
            xls.SetCellValue(40, 5, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(40, 6, 1);
            xls.SetCellValue(40, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(40, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(40, 8, xls.AddFormat(fmt));
            xls.SetCellValue(40, 8, new TFormula("=(    IF(E40<>1,VLOOKUP(E40,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)   )     "
            + " *       IF(F40<>1,VLOOKUP(F40,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) * IF(G40<>1,VLOOKUP(G40,'Gral"
            + " Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(40, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(40, 9, xls.AddFormat(fmt));
            xls.SetCellValue(40, 9, new TFormula("=C40*H40"));

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

            fmt = xls.GetCellVisibleFormatDef(41, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(41, 9, xls.AddFormat(fmt));

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
            xls.SetCellValue(42, 3, 47.4285714285714);
            xls.SetCellValue(42, 5, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(42, 6, 1);
            xls.SetCellValue(42, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(42, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(42, 8, xls.AddFormat(fmt));
            xls.SetCellValue(42, 8, new TFormula("=(    IF(E42<>1,VLOOKUP(E42,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)   )     "
            + " *       IF(F42<>1,VLOOKUP(F42,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) * IF(G42<>1,VLOOKUP(G42,'Gral"
            + " Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(42, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(42, 9, xls.AddFormat(fmt));
            xls.SetCellValue(42, 9, new TFormula("=C42*H42"));

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
            xls.SetCellValue(43, 3, 5.32);
            xls.SetCellValue(43, 5, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(43, 6, 1);
            xls.SetCellValue(43, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(43, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(43, 8, xls.AddFormat(fmt));
            xls.SetCellValue(43, 8, new TFormula("=(    IF(E43<>1,VLOOKUP(E43,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)   )     "
            + " *       IF(F43<>1,VLOOKUP(F43,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) * IF(G43<>1,VLOOKUP(G43,'Gral"
            + " Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(43, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(43, 9, xls.AddFormat(fmt));
            xls.SetCellValue(43, 9, new TFormula("=C43*H43"));

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
            xls.SetCellValue(44, 3, 0.24);
            xls.SetCellValue(44, 5, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(44, 6, 1);
            xls.SetCellValue(44, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(44, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(44, 8, xls.AddFormat(fmt));
            xls.SetCellValue(44, 8, new TFormula("=(    IF(E44<>1,VLOOKUP(E44,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)   )     "
            + " *       IF(F44<>1,VLOOKUP(F44,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) * IF(G44<>1,VLOOKUP(G44,'Gral"
            + " Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(44, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(44, 9, xls.AddFormat(fmt));
            xls.SetCellValue(44, 9, new TFormula("=C44*H44"));

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
            xls.SetCellValue(45, 3, 6.2);
            xls.SetCellValue(45, 5, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(45, 6, 1);
            xls.SetCellValue(45, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(45, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(45, 8, xls.AddFormat(fmt));
            xls.SetCellValue(45, 8, new TFormula("=(    IF(E45<>1,VLOOKUP(E45,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)   )     "
            + " *       IF(F45<>1,VLOOKUP(F45,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) * IF(G45<>1,VLOOKUP(G45,'Gral"
            + " Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(45, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(45, 9, xls.AddFormat(fmt));
            xls.SetCellValue(45, 9, new TFormula("=C45*H45"));

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
            xls.SetCellValue(46, 3, 2.1);
            xls.SetCellValue(46, 5, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(46, 6, 1);
            xls.SetCellValue(46, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(46, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(46, 8, xls.AddFormat(fmt));
            xls.SetCellValue(46, 8, new TFormula("=(    IF(E46<>1,VLOOKUP(E46,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)   )     "
            + " *       IF(F46<>1,VLOOKUP(F46,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) * IF(G46<>1,VLOOKUP(G46,'Gral"
            + " Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(46, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(46, 9, xls.AddFormat(fmt));
            xls.SetCellValue(46, 9, new TFormula("=C46*H46"));

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

            fmt = xls.GetCellVisibleFormatDef(47, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(47, 9, xls.AddFormat(fmt));

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

            fmt = xls.GetCellVisibleFormatDef(48, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(48, 9, xls.AddFormat(fmt));

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
            xls.SetCellValue(49, 3, 40.34);
            xls.SetCellValue(49, 5, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(49, 6, 1);
            xls.SetCellValue(49, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(49, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(49, 8, xls.AddFormat(fmt));
            xls.SetCellValue(49, 8, new TFormula("=(    IF(E49<>1,VLOOKUP(E49,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)   )     "
            + " *       IF(F49<>1,VLOOKUP(F49,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) * IF(G49<>1,VLOOKUP(G49,'Gral"
            + " Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(49, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(49, 9, xls.AddFormat(fmt));
            xls.SetCellValue(49, 9, new TFormula("=C49*H49"));

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
            xls.SetCellValue(50, 3, 0.04);
            xls.SetCellValue(50, 5, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(50, 6, 1);
            xls.SetCellValue(50, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(50, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(50, 8, xls.AddFormat(fmt));
            xls.SetCellValue(50, 8, new TFormula("=(    IF(E50<>1,VLOOKUP(E50,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)   )     "
            + " *       IF(F50<>1,VLOOKUP(F50,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) * IF(G50<>1,VLOOKUP(G50,'Gral"
            + " Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(50, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(50, 9, xls.AddFormat(fmt));
            xls.SetCellValue(50, 9, new TFormula("=C50*H50"));

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
            xls.SetCellValue(51, 3, 5.75);
            xls.SetCellValue(51, 5, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(51, 6, 1);
            xls.SetCellValue(51, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(51, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(51, 8, xls.AddFormat(fmt));
            xls.SetCellValue(51, 8, new TFormula("=(    IF(E51<>1,VLOOKUP(E51,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)   )     "
            + " *       IF(F51<>1,VLOOKUP(F51,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) * IF(G51<>1,VLOOKUP(G51,'Gral"
            + " Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(51, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(51, 9, xls.AddFormat(fmt));
            xls.SetCellValue(51, 9, new TFormula("=C51*H51"));

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
            xls.SetCellValue(52, 3, 0);
            xls.SetCellValue(52, 5, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(52, 6, 1);
            xls.SetCellValue(52, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(52, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(52, 8, xls.AddFormat(fmt));
            xls.SetCellValue(52, 8, new TFormula("=(    IF(E52<>1,VLOOKUP(E52,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)   )     "
            + " *       IF(F52<>1,VLOOKUP(F52,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) * IF(G52<>1,VLOOKUP(G52,'Gral"
            + " Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(52, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(52, 9, xls.AddFormat(fmt));
            xls.SetCellValue(52, 9, new TFormula("=C52*H52"));

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
            xls.SetCellValue(53, 3, 3.4);
            xls.SetCellValue(53, 5, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(53, 6, 1);
            xls.SetCellValue(53, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(53, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(53, 8, xls.AddFormat(fmt));
            xls.SetCellValue(53, 8, new TFormula("=(    IF(E53<>1,VLOOKUP(E53,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)   )     "
            + " *       IF(F53<>1,VLOOKUP(F53,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) * IF(G53<>1,VLOOKUP(G53,'Gral"
            + " Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(53, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(53, 9, xls.AddFormat(fmt));
            xls.SetCellValue(53, 9, new TFormula("=C53*H53"));

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
            xls.SetCellValue(54, 3, 4);
            xls.SetCellValue(54, 5, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(54, 6, 1);
            xls.SetCellValue(54, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(54, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(54, 8, xls.AddFormat(fmt));
            xls.SetCellValue(54, 8, new TFormula("=(    IF(E54<>1,VLOOKUP(E54,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)   )     "
            + " *       IF(F54<>1,VLOOKUP(F54,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) * IF(G54<>1,VLOOKUP(G54,'Gral"
            + " Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(54, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(54, 9, xls.AddFormat(fmt));
            xls.SetCellValue(54, 9, new TFormula("=C54*H54"));

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
            xls.SetCellValue(55, 3, 13);
            xls.SetCellValue(55, 5, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(55, 6, 1);
            xls.SetCellValue(55, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(55, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(55, 8, xls.AddFormat(fmt));
            xls.SetCellValue(55, 8, new TFormula("=(    IF(E55<>1,VLOOKUP(E55,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)   )     "
            + " *       IF(F55<>1,VLOOKUP(F55,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) * IF(G55<>1,VLOOKUP(G55,'Gral"
            + " Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(55, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(55, 9, xls.AddFormat(fmt));
            xls.SetCellValue(55, 9, new TFormula("=C55*H55"));

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
            xls.SetCellValue(56, 3, 0.3);
            xls.SetCellValue(56, 5, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(56, 6, 1);
            xls.SetCellValue(56, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(56, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(56, 8, xls.AddFormat(fmt));
            xls.SetCellValue(56, 8, new TFormula("=(    IF(E56<>1,VLOOKUP(E56,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)   )     "
            + " *       IF(F56<>1,VLOOKUP(F56,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) * IF(G56<>1,VLOOKUP(G56,'Gral"
            + " Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(56, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(56, 9, xls.AddFormat(fmt));
            xls.SetCellValue(56, 9, new TFormula("=C56*H56"));

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
            xls.SetCellValue(57, 3, 8.9);
            xls.SetCellValue(57, 5, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(57, 6, 1);
            xls.SetCellValue(57, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(57, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(57, 8, xls.AddFormat(fmt));
            xls.SetCellValue(57, 8, new TFormula("=(    IF(E57<>1,VLOOKUP(E57,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)   )     "
            + " *       IF(F57<>1,VLOOKUP(F57,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) * IF(G57<>1,VLOOKUP(G57,'Gral"
            + " Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(57, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(57, 9, xls.AddFormat(fmt));
            xls.SetCellValue(57, 9, new TFormula("=C57*H57"));

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
            xls.SetCellValue(58, 3, 7.84);
            xls.SetCellValue(58, 5, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(58, 6, 1);
            xls.SetCellValue(58, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(58, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(58, 8, xls.AddFormat(fmt));
            xls.SetCellValue(58, 8, new TFormula("=(    IF(E58<>1,VLOOKUP(E58,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)   )     "
            + " *       IF(F58<>1,VLOOKUP(F58,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) * IF(G58<>1,VLOOKUP(G58,'Gral"
            + " Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(58, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(58, 9, xls.AddFormat(fmt));
            xls.SetCellValue(58, 9, new TFormula("=C58*H58"));

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

            fmt = xls.GetCellVisibleFormatDef(59, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(59, 9, xls.AddFormat(fmt));

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
            xls.SetCellValue(60, 3, 25);
            xls.SetCellValue(60, 5, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(60, 6, 1);
            xls.SetCellValue(60, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(60, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(60, 8, xls.AddFormat(fmt));
            xls.SetCellValue(60, 8, new TFormula("=(    IF(E60<>1,VLOOKUP(E60,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)   )     "
            + " *       IF(F60<>1,VLOOKUP(F60,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) * IF(G60<>1,VLOOKUP(G60,'Gral"
            + " Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(60, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(60, 9, xls.AddFormat(fmt));
            xls.SetCellValue(60, 9, new TFormula("=C60*H60"));

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
            xls.SetCellValue(61, 3, 0);
            xls.SetCellValue(61, 5, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(61, 6, 1);
            xls.SetCellValue(61, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(61, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(61, 8, xls.AddFormat(fmt));
            xls.SetCellValue(61, 8, new TFormula("=(    IF(E61<>1,VLOOKUP(E61,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)   )     "
            + " *       IF(F61<>1,VLOOKUP(F61,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) * IF(G61<>1,VLOOKUP(G61,'Gral"
            + " Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(61, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(61, 9, xls.AddFormat(fmt));
            xls.SetCellValue(61, 9, new TFormula("=C61*H61"));

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

            fmt = xls.GetCellVisibleFormatDef(62, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(62, 9, xls.AddFormat(fmt));

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
            xls.SetCellValue(63, 3, 3);
            xls.SetCellValue(63, 5, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(63, 6, 1);
            xls.SetCellValue(63, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(63, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(63, 8, xls.AddFormat(fmt));
            xls.SetCellValue(63, 8, new TFormula("=(    IF(E63<>1,VLOOKUP(E63,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)   )     "
            + " *       IF(F63<>1,VLOOKUP(F63,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) * IF(G63<>1,VLOOKUP(G63,'Gral"
            + " Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(63, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(63, 9, xls.AddFormat(fmt));
            xls.SetCellValue(63, 9, new TFormula("=C63*H63"));

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
            xls.SetCellValue(64, 3, 3);
            xls.SetCellValue(64, 5, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(64, 6, 1);
            xls.SetCellValue(64, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(64, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(64, 8, xls.AddFormat(fmt));
            xls.SetCellValue(64, 8, new TFormula("=(    IF(E64<>1,VLOOKUP(E64,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)   )     "
            + " *       IF(F64<>1,VLOOKUP(F64,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) * IF(G64<>1,VLOOKUP(G64,'Gral"
            + " Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(64, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(64, 9, xls.AddFormat(fmt));
            xls.SetCellValue(64, 9, new TFormula("=C64*H64"));

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
            xls.SetCellValue(65, 3, 5.8);
            xls.SetCellValue(65, 5, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(65, 6, 1);
            xls.SetCellValue(65, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(65, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(65, 8, xls.AddFormat(fmt));
            xls.SetCellValue(65, 8, new TFormula("=(    IF(E65<>1,VLOOKUP(E65,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)   )     "
            + " *       IF(F65<>1,VLOOKUP(F65,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) * IF(G65<>1,VLOOKUP(G65,'Gral"
            + " Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(65, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(65, 9, xls.AddFormat(fmt));
            xls.SetCellValue(65, 9, new TFormula("=C65*H65"));

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
            xls.SetCellValue(66, 3, 1.2);
            xls.SetCellValue(66, 5, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(66, 6, 1);
            xls.SetCellValue(66, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(66, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(66, 8, xls.AddFormat(fmt));
            xls.SetCellValue(66, 8, new TFormula("=(    IF(E66<>1,VLOOKUP(E66,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)   )     "
            + " *       IF(F66<>1,VLOOKUP(F66,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) * IF(G66<>1,VLOOKUP(G66,'Gral"
            + " Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(66, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(66, 9, xls.AddFormat(fmt));
            xls.SetCellValue(66, 9, new TFormula("=C66*H66"));

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
            xls.SetCellValue(67, 3, 1.8);
            xls.SetCellValue(67, 5, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(67, 6, 1);
            xls.SetCellValue(67, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(67, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(67, 8, xls.AddFormat(fmt));
            xls.SetCellValue(67, 8, new TFormula("=(    IF(E67<>1,VLOOKUP(E67,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)   )     "
            + " *       IF(F67<>1,VLOOKUP(F67,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) * IF(G67<>1,VLOOKUP(G67,'Gral"
            + " Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(67, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(67, 9, xls.AddFormat(fmt));
            xls.SetCellValue(67, 9, new TFormula("=C67*H67"));

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
            xls.SetCellValue(68, 3, 1);
            xls.SetCellValue(68, 5, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(68, 6, 1);
            xls.SetCellValue(68, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(68, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(68, 8, xls.AddFormat(fmt));
            xls.SetCellValue(68, 8, new TFormula("=(    IF(E68<>1,VLOOKUP(E68,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)   )     "
            + " *       IF(F68<>1,VLOOKUP(F68,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) * IF(G68<>1,VLOOKUP(G68,'Gral"
            + " Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(68, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(68, 9, xls.AddFormat(fmt));
            xls.SetCellValue(68, 9, new TFormula("=C68*H68"));

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
            xls.SetCellValue(69, 3, 0.28);
            xls.SetCellValue(69, 5, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(69, 6, 1);
            xls.SetCellValue(69, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(69, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(69, 8, xls.AddFormat(fmt));
            xls.SetCellValue(69, 8, new TFormula("=(    IF(E69<>1,VLOOKUP(E69,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)   )     "
            + " *       IF(F69<>1,VLOOKUP(F69,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) * IF(G69<>1,VLOOKUP(G69,'Gral"
            + " Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(69, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(69, 9, xls.AddFormat(fmt));
            xls.SetCellValue(69, 9, new TFormula("=C69*H69"));

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
            xls.SetCellValue(70, 3, 1.9);
            xls.SetCellValue(70, 5, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(70, 6, 1);
            xls.SetCellValue(70, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(70, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(70, 8, xls.AddFormat(fmt));
            xls.SetCellValue(70, 8, new TFormula("=(    IF(E70<>1,VLOOKUP(E70,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)   )     "
            + " *       IF(F70<>1,VLOOKUP(F70,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) * IF(G70<>1,VLOOKUP(G70,'Gral"
            + " Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(70, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(70, 9, xls.AddFormat(fmt));
            xls.SetCellValue(70, 9, new TFormula("=C70*H70"));

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
            xls.SetCellValue(71, 3, 0.1);
            xls.SetCellValue(71, 5, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(71, 6, 1);
            xls.SetCellValue(71, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(71, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(71, 8, xls.AddFormat(fmt));
            xls.SetCellValue(71, 8, new TFormula("=(    IF(E71<>1,VLOOKUP(E71,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)   )     "
            + " *       IF(F71<>1,VLOOKUP(F71,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) * IF(G71<>1,VLOOKUP(G71,'Gral"
            + " Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(71, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(71, 9, xls.AddFormat(fmt));
            xls.SetCellValue(71, 9, new TFormula("=C71*H71"));

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

            fmt = xls.GetCellVisibleFormatDef(72, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(72, 9, xls.AddFormat(fmt));

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
            xls.SetCellValue(73, 3, 31);
            xls.SetCellValue(73, 5, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(73, 6, 1);
            xls.SetCellValue(73, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(73, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(73, 8, xls.AddFormat(fmt));
            xls.SetCellValue(73, 8, new TFormula("=(    IF(E73<>1,VLOOKUP(E73,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)   )     "
            + " *       IF(F73<>1,VLOOKUP(F73,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) * IF(G73<>1,VLOOKUP(G73,'Gral"
            + " Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(73, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(73, 9, xls.AddFormat(fmt));
            xls.SetCellValue(73, 9, new TFormula("=C73*H73"));

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
            xls.SetCellValue(74, 3, 0.04);
            xls.SetCellValue(74, 5, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(74, 6, 1);
            xls.SetCellValue(74, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(74, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(74, 8, xls.AddFormat(fmt));
            xls.SetCellValue(74, 8, new TFormula("=(    IF(E74<>1,VLOOKUP(E74,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)   )     "
            + " *       IF(F74<>1,VLOOKUP(F74,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) * IF(G74<>1,VLOOKUP(G74,'Gral"
            + " Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(74, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(74, 9, xls.AddFormat(fmt));
            xls.SetCellValue(74, 9, new TFormula("=C74*H74"));

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
            xls.SetCellValue(75, 3, 5.5);
            xls.SetCellValue(75, 5, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(75, 6, 1);
            xls.SetCellValue(75, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(75, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(75, 8, xls.AddFormat(fmt));
            xls.SetCellValue(75, 8, new TFormula("=(    IF(E75<>1,VLOOKUP(E75,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)   )     "
            + " *       IF(F75<>1,VLOOKUP(F75,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) * IF(G75<>1,VLOOKUP(G75,'Gral"
            + " Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(75, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(75, 9, xls.AddFormat(fmt));
            xls.SetCellValue(75, 9, new TFormula("=C75*H75"));

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
            xls.SetCellValue(76, 3, 0);
            xls.SetCellValue(76, 5, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(76, 6, 1);
            xls.SetCellValue(76, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(76, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(76, 8, xls.AddFormat(fmt));
            xls.SetCellValue(76, 8, new TFormula("=(    IF(E76<>1,VLOOKUP(E76,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)   )     "
            + " *       IF(F76<>1,VLOOKUP(F76,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) * IF(G76<>1,VLOOKUP(G76,'Gral"
            + " Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(76, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(76, 9, xls.AddFormat(fmt));
            xls.SetCellValue(76, 9, new TFormula("=C76*H76"));

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
            xls.SetCellValue(77, 3, 3.4);
            xls.SetCellValue(77, 5, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(77, 6, 1);
            xls.SetCellValue(77, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(77, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(77, 8, xls.AddFormat(fmt));
            xls.SetCellValue(77, 8, new TFormula("=(    IF(E77<>1,VLOOKUP(E77,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)   )     "
            + " *       IF(F77<>1,VLOOKUP(F77,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) * IF(G77<>1,VLOOKUP(G77,'Gral"
            + " Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(77, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(77, 9, xls.AddFormat(fmt));
            xls.SetCellValue(77, 9, new TFormula("=C77*H77"));

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
            xls.SetCellValue(78, 3, 2.5);
            xls.SetCellValue(78, 5, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(78, 6, 1);
            xls.SetCellValue(78, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(78, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(78, 8, xls.AddFormat(fmt));
            xls.SetCellValue(78, 8, new TFormula("=(    IF(E78<>1,VLOOKUP(E78,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)   )     "
            + " *       IF(F78<>1,VLOOKUP(F78,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) * IF(G78<>1,VLOOKUP(G78,'Gral"
            + " Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(78, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(78, 9, xls.AddFormat(fmt));
            xls.SetCellValue(78, 9, new TFormula("=C78*H78"));

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
            xls.SetCellValue(79, 3, 11.7);
            xls.SetCellValue(79, 5, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(79, 6, 1);
            xls.SetCellValue(79, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(79, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(79, 8, xls.AddFormat(fmt));
            xls.SetCellValue(79, 8, new TFormula("=(    IF(E79<>1,VLOOKUP(E79,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)   )     "
            + " *       IF(F79<>1,VLOOKUP(F79,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) * IF(G79<>1,VLOOKUP(G79,'Gral"
            + " Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(79, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(79, 9, xls.AddFormat(fmt));
            xls.SetCellValue(79, 9, new TFormula("=C79*H79"));

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
            xls.SetCellValue(80, 3, 0.36);
            xls.SetCellValue(80, 5, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(80, 6, 1);
            xls.SetCellValue(80, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(80, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(80, 8, xls.AddFormat(fmt));
            xls.SetCellValue(80, 8, new TFormula("=(    IF(E80<>1,VLOOKUP(E80,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)   )     "
            + " *       IF(F80<>1,VLOOKUP(F80,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) * IF(G80<>1,VLOOKUP(G80,'Gral"
            + " Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(80, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(80, 9, xls.AddFormat(fmt));
            xls.SetCellValue(80, 9, new TFormula("=C80*H80"));

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
            xls.SetCellValue(81, 3, 3.91);
            xls.SetCellValue(81, 5, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(81, 6, 1);
            xls.SetCellValue(81, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(81, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(81, 8, xls.AddFormat(fmt));
            xls.SetCellValue(81, 8, new TFormula("=(    IF(E81<>1,VLOOKUP(E81,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)   )     "
            + " *       IF(F81<>1,VLOOKUP(F81,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) * IF(G81<>1,VLOOKUP(G81,'Gral"
            + " Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(81, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(81, 9, xls.AddFormat(fmt));
            xls.SetCellValue(81, 9, new TFormula("=C81*H81"));

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
            xls.SetCellValue(82, 3, 7.36);
            xls.SetCellValue(82, 5, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(82, 6, 1);
            xls.SetCellValue(82, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(82, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(82, 8, xls.AddFormat(fmt));
            xls.SetCellValue(82, 8, new TFormula("=(    IF(E82<>1,VLOOKUP(E82,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)   )     "
            + " *       IF(F82<>1,VLOOKUP(F82,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) * IF(G82<>1,VLOOKUP(G82,'Gral"
            + " Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(82, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(82, 9, xls.AddFormat(fmt));
            xls.SetCellValue(82, 9, new TFormula("=C82*H82"));

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

            fmt = xls.GetCellVisibleFormatDef(83, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(83, 9, xls.AddFormat(fmt));

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
            xls.SetCellValue(84, 3, 65);
            xls.SetCellValue(84, 5, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(84, 6, 1);
            xls.SetCellValue(84, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(84, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(84, 8, xls.AddFormat(fmt));
            xls.SetCellValue(84, 8, new TFormula("=(    IF(E84<>1,VLOOKUP(E84,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)   )     "
            + " *       IF(F84<>1,VLOOKUP(F84,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) * IF(G84<>1,VLOOKUP(G84,'Gral"
            + " Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(84, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(84, 9, xls.AddFormat(fmt));
            xls.SetCellValue(84, 9, new TFormula("=C84*H84"));

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
            xls.SetCellValue(85, 3, 0);
            xls.SetCellValue(85, 5, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(85, 6, 1);
            xls.SetCellValue(85, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(85, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(85, 8, xls.AddFormat(fmt));
            xls.SetCellValue(85, 8, new TFormula("=(    IF(E85<>1,VLOOKUP(E85,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)   )     "
            + " *       IF(F85<>1,VLOOKUP(F85,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) * IF(G85<>1,VLOOKUP(G85,'Gral"
            + " Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(85, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(85, 9, xls.AddFormat(fmt));
            xls.SetCellValue(85, 9, new TFormula("=C85*H85"));

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

            fmt = xls.GetCellVisibleFormatDef(86, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(86, 9, xls.AddFormat(fmt));

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
            xls.SetCellValue(87, 3, 6.5);
            xls.SetCellValue(87, 5, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(87, 6, 1);
            xls.SetCellValue(87, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(87, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(87, 8, xls.AddFormat(fmt));
            xls.SetCellValue(87, 8, new TFormula("=(    IF(E87<>1,VLOOKUP(E87,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)   )     "
            + " *       IF(F87<>1,VLOOKUP(F87,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) * IF(G87<>1,VLOOKUP(G87,'Gral"
            + " Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(87, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(87, 9, xls.AddFormat(fmt));
            xls.SetCellValue(87, 9, new TFormula("=C87*H87"));

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
            xls.SetCellValue(88, 3, 6);
            xls.SetCellValue(88, 5, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(88, 6, 1);
            xls.SetCellValue(88, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(88, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(88, 8, xls.AddFormat(fmt));
            xls.SetCellValue(88, 8, new TFormula("=(    IF(E88<>1,VLOOKUP(E88,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)   )     "
            + " *       IF(F88<>1,VLOOKUP(F88,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) * IF(G88<>1,VLOOKUP(G88,'Gral"
            + " Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(88, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(88, 9, xls.AddFormat(fmt));
            xls.SetCellValue(88, 9, new TFormula("=C88*H88"));

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
            xls.SetCellValue(89, 3, 8.5);
            xls.SetCellValue(89, 5, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(89, 6, 1);
            xls.SetCellValue(89, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(89, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(89, 8, xls.AddFormat(fmt));
            xls.SetCellValue(89, 8, new TFormula("=(    IF(E89<>1,VLOOKUP(E89,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)   )     "
            + " *       IF(F89<>1,VLOOKUP(F89,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) * IF(G89<>1,VLOOKUP(G89,'Gral"
            + " Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(89, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(89, 9, xls.AddFormat(fmt));
            xls.SetCellValue(89, 9, new TFormula("=C89*H89"));

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
            xls.SetCellValue(90, 3, 2.13);
            xls.SetCellValue(90, 5, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(90, 6, 1);
            xls.SetCellValue(90, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(90, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(90, 8, xls.AddFormat(fmt));
            xls.SetCellValue(90, 8, new TFormula("=(    IF(E90<>1,VLOOKUP(E90,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)   )     "
            + " *       IF(F90<>1,VLOOKUP(F90,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) * IF(G90<>1,VLOOKUP(G90,'Gral"
            + " Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(90, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(90, 9, xls.AddFormat(fmt));
            xls.SetCellValue(90, 9, new TFormula("=C90*H90"));

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
            xls.SetCellValue(91, 3, 4.8);
            xls.SetCellValue(91, 5, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(91, 6, 1);
            xls.SetCellValue(91, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(91, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(91, 8, xls.AddFormat(fmt));
            xls.SetCellValue(91, 8, new TFormula("=(    IF(E91<>1,VLOOKUP(E91,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)   )     "
            + " *       IF(F91<>1,VLOOKUP(F91,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) * IF(G91<>1,VLOOKUP(G91,'Gral"
            + " Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(91, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(91, 9, xls.AddFormat(fmt));
            xls.SetCellValue(91, 9, new TFormula("=C91*H91"));

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
            xls.SetCellValue(92, 3, 2.3);
            xls.SetCellValue(92, 5, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(92, 6, 1);
            xls.SetCellValue(92, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(92, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(92, 8, xls.AddFormat(fmt));
            xls.SetCellValue(92, 8, new TFormula("=(    IF(E92<>1,VLOOKUP(E92,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)   )     "
            + " *       IF(F92<>1,VLOOKUP(F92,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) * IF(G92<>1,VLOOKUP(G92,'Gral"
            + " Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(92, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(92, 9, xls.AddFormat(fmt));
            xls.SetCellValue(92, 9, new TFormula("=C92*H92"));

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
            xls.SetCellValue(93, 3, 0.43);
            xls.SetCellValue(93, 5, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(93, 6, 1);
            xls.SetCellValue(93, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(93, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(93, 8, xls.AddFormat(fmt));
            xls.SetCellValue(93, 8, new TFormula("=(    IF(E93<>1,VLOOKUP(E93,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)   )     "
            + " *       IF(F93<>1,VLOOKUP(F93,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) * IF(G93<>1,VLOOKUP(G93,'Gral"
            + " Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(93, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(93, 9, xls.AddFormat(fmt));
            xls.SetCellValue(93, 9, new TFormula("=C93*H93"));

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
            xls.SetCellValue(94, 3, 3);
            xls.SetCellValue(94, 5, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(94, 6, 1);
            xls.SetCellValue(94, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(94, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(94, 8, xls.AddFormat(fmt));
            xls.SetCellValue(94, 8, new TFormula("=(    IF(E94<>1,VLOOKUP(E94,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)   )     "
            + " *       IF(F94<>1,VLOOKUP(F94,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) * IF(G94<>1,VLOOKUP(G94,'Gral"
            + " Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(94, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(94, 9, xls.AddFormat(fmt));
            xls.SetCellValue(94, 9, new TFormula("=C94*H94"));

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
            xls.SetCellValue(95, 3, 0.1);
            xls.SetCellValue(95, 5, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(95, 6, 1);
            xls.SetCellValue(95, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(95, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(95, 8, xls.AddFormat(fmt));
            xls.SetCellValue(95, 8, new TFormula("=(    IF(E95<>1,VLOOKUP(E95,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)   )     "
            + " *       IF(F95<>1,VLOOKUP(F95,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) * IF(G95<>1,VLOOKUP(G95,'Gral"
            + " Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(95, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(95, 9, xls.AddFormat(fmt));
            xls.SetCellValue(95, 9, new TFormula("=C95*H95"));

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

            fmt = xls.GetCellVisibleFormatDef(96, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            xls.SetCellFormat(96, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(96, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            xls.SetCellFormat(96, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(96, 7);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            xls.SetCellFormat(96, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(96, 8);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(96, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(96, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(96, 9, xls.AddFormat(fmt));

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
            xls.SetCellValue(97, 3, 28);
            xls.SetCellValue(97, 5, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(97, 6, 1);
            xls.SetCellValue(97, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(97, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(97, 8, xls.AddFormat(fmt));
            xls.SetCellValue(97, 8, new TFormula("=(    IF(E97<>1,VLOOKUP(E97,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)   )     "
            + " *       IF(F97<>1,VLOOKUP(F97,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) * IF(G97<>1,VLOOKUP(G97,'Gral"
            + " Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(97, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(97, 9, xls.AddFormat(fmt));
            xls.SetCellValue(97, 9, new TFormula("=C97*H97"));

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
            xls.SetCellValue(98, 3, 0.04);
            xls.SetCellValue(98, 5, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(98, 6, 1);
            xls.SetCellValue(98, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(98, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(98, 8, xls.AddFormat(fmt));
            xls.SetCellValue(98, 8, new TFormula("=(    IF(E98<>1,VLOOKUP(E98,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)   )     "
            + " *       IF(F98<>1,VLOOKUP(F98,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) * IF(G98<>1,VLOOKUP(G98,'Gral"
            + " Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(98, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(98, 9, xls.AddFormat(fmt));
            xls.SetCellValue(98, 9, new TFormula("=C98*H98"));

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
            xls.SetCellValue(99, 3, 5.78);
            xls.SetCellValue(99, 5, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(99, 6, 1);
            xls.SetCellValue(99, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(99, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(99, 8, xls.AddFormat(fmt));
            xls.SetCellValue(99, 8, new TFormula("=(    IF(E99<>1,VLOOKUP(E99,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)   )     "
            + " *       IF(F99<>1,VLOOKUP(F99,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) * IF(G99<>1,VLOOKUP(G99,'Gral"
            + " Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(99, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(99, 9, xls.AddFormat(fmt));
            xls.SetCellValue(99, 9, new TFormula("=C99*H99"));

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
            xls.SetCellValue(100, 3, 0);
            xls.SetCellValue(100, 5, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(100, 6, 1);
            xls.SetCellValue(100, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(100, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(100, 8, xls.AddFormat(fmt));
            xls.SetCellValue(100, 8, new TFormula("=(    IF(E100<>1,VLOOKUP(E100,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)   )   "
            + "   *       IF(F100<>1,VLOOKUP(F100,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) *"
            + " IF(G100<>1,VLOOKUP(G100,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(100, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(100, 9, xls.AddFormat(fmt));
            xls.SetCellValue(100, 9, new TFormula("=C100*H100"));

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
            xls.SetCellValue(101, 3, 3.71);
            xls.SetCellValue(101, 5, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(101, 6, 1);
            xls.SetCellValue(101, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(101, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(101, 8, xls.AddFormat(fmt));
            xls.SetCellValue(101, 8, new TFormula("=(    IF(E101<>1,VLOOKUP(E101,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)   )   "
            + "   *       IF(F101<>1,VLOOKUP(F101,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) *"
            + " IF(G101<>1,VLOOKUP(G101,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(101, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(101, 9, xls.AddFormat(fmt));
            xls.SetCellValue(101, 9, new TFormula("=C101*H101"));

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
            xls.SetCellValue(102, 3, 2.2);
            xls.SetCellValue(102, 5, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(102, 6, 1);
            xls.SetCellValue(102, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(102, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(102, 8, xls.AddFormat(fmt));
            xls.SetCellValue(102, 8, new TFormula("=(    IF(E102<>1,VLOOKUP(E102,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)   )   "
            + "   *       IF(F102<>1,VLOOKUP(F102,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) *"
            + " IF(G102<>1,VLOOKUP(G102,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(102, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(102, 9, xls.AddFormat(fmt));
            xls.SetCellValue(102, 9, new TFormula("=C102*H102"));

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
            xls.SetCellValue(103, 3, 12.2);
            xls.SetCellValue(103, 5, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(103, 6, 1);
            xls.SetCellValue(103, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(103, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(103, 8, xls.AddFormat(fmt));
            xls.SetCellValue(103, 8, new TFormula("=(    IF(E103<>1,VLOOKUP(E103,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)   )   "
            + "   *       IF(F103<>1,VLOOKUP(F103,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) *"
            + " IF(G103<>1,VLOOKUP(G103,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(103, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(103, 9, xls.AddFormat(fmt));
            xls.SetCellValue(103, 9, new TFormula("=C103*H103"));

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
            xls.SetCellValue(104, 3, 0.36);
            xls.SetCellValue(104, 5, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(104, 6, 1);
            xls.SetCellValue(104, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(104, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(104, 8, xls.AddFormat(fmt));
            xls.SetCellValue(104, 8, new TFormula("=(    IF(E104<>1,VLOOKUP(E104,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)   )   "
            + "   *       IF(F104<>1,VLOOKUP(F104,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) *"
            + " IF(G104<>1,VLOOKUP(G104,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(104, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(104, 9, xls.AddFormat(fmt));
            xls.SetCellValue(104, 9, new TFormula("=C104*H104"));

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
            xls.SetCellValue(105, 3, 4.54);
            xls.SetCellValue(105, 5, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(105, 6, 1);
            xls.SetCellValue(105, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(105, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(105, 8, xls.AddFormat(fmt));
            xls.SetCellValue(105, 8, new TFormula("=(    IF(E105<>1,VLOOKUP(E105,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)   )   "
            + "   *       IF(F105<>1,VLOOKUP(F105,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) *"
            + " IF(G105<>1,VLOOKUP(G105,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(105, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(105, 9, xls.AddFormat(fmt));
            xls.SetCellValue(105, 9, new TFormula("=C105*H105"));

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
            xls.SetCellValue(106, 3, 7.91);
            xls.SetCellValue(106, 5, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(106, 6, 1);
            xls.SetCellValue(106, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(106, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(106, 8, xls.AddFormat(fmt));
            xls.SetCellValue(106, 8, new TFormula("=(    IF(E106<>1,VLOOKUP(E106,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)   )   "
            + "   *       IF(F106<>1,VLOOKUP(F106,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) *"
            + " IF(G106<>1,VLOOKUP(G106,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(106, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(106, 9, xls.AddFormat(fmt));
            xls.SetCellValue(106, 9, new TFormula("=C106*H106"));

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

            fmt = xls.GetCellVisibleFormatDef(107, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            xls.SetCellFormat(107, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(107, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            xls.SetCellFormat(107, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(107, 7);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            xls.SetCellFormat(107, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(107, 8);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(107, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(107, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(107, 9, xls.AddFormat(fmt));

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
            xls.SetCellValue(108, 3, 53);
            xls.SetCellValue(108, 5, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(108, 6, 1);
            xls.SetCellValue(108, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(108, 8);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(108, 8, xls.AddFormat(fmt));
            xls.SetCellValue(108, 8, 0.705019741);

            fmt = xls.GetCellVisibleFormatDef(108, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(108, 9, xls.AddFormat(fmt));
            xls.SetCellValue(108, 9, new TFormula("=C108*H108"));

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
            xls.SetCellValue(109, 3, 0);
            xls.SetCellValue(109, 5, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(109, 6, 1);
            xls.SetCellValue(109, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(109, 8);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(109, 8, xls.AddFormat(fmt));
            xls.SetCellValue(109, 8, 0.705019741);

            fmt = xls.GetCellVisibleFormatDef(109, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(109, 9, xls.AddFormat(fmt));
            xls.SetCellValue(109, 9, new TFormula("=C109*H109"));

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

            fmt = xls.GetCellVisibleFormatDef(110, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(110, 9, xls.AddFormat(fmt));

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
            xls.SetCellValue(111, 3, 4.6);
            xls.SetCellValue(111, 5, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(111, 6, 1);
            xls.SetCellValue(111, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(111, 8);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(111, 8, xls.AddFormat(fmt));
            xls.SetCellValue(111, 8, 0.705019741);

            fmt = xls.GetCellVisibleFormatDef(111, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(111, 9, xls.AddFormat(fmt));
            xls.SetCellValue(111, 9, new TFormula("=C111*H111"));

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
            xls.SetCellValue(112, 3, 2.3);
            xls.SetCellValue(112, 5, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(112, 6, 1);
            xls.SetCellValue(112, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(112, 8);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(112, 8, xls.AddFormat(fmt));
            xls.SetCellValue(112, 8, 0.705019741);

            fmt = xls.GetCellVisibleFormatDef(112, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(112, 9, xls.AddFormat(fmt));
            xls.SetCellValue(112, 9, new TFormula("=C112*H112"));

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
            xls.SetCellValue(113, 3, 1.2);
            xls.SetCellValue(113, 5, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(113, 6, 1);
            xls.SetCellValue(113, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(113, 8);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(113, 8, xls.AddFormat(fmt));
            xls.SetCellValue(113, 8, 0.705019741);

            fmt = xls.GetCellVisibleFormatDef(113, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(113, 9, xls.AddFormat(fmt));
            xls.SetCellValue(113, 9, new TFormula("=C113*H113"));

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
            xls.SetCellValue(114, 3, 0.83);
            xls.SetCellValue(114, 5, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(114, 6, 1);
            xls.SetCellValue(114, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(114, 8);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(114, 8, xls.AddFormat(fmt));
            xls.SetCellValue(114, 8, 0.705019741);

            fmt = xls.GetCellVisibleFormatDef(114, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(114, 9, xls.AddFormat(fmt));
            xls.SetCellValue(114, 9, new TFormula("=C114*H114"));

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
            xls.SetCellValue(115, 3, 0);
            xls.SetCellValue(115, 5, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(115, 6, 1);
            xls.SetCellValue(115, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(115, 8);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(115, 8, xls.AddFormat(fmt));
            xls.SetCellValue(115, 8, 0.705019741);

            fmt = xls.GetCellVisibleFormatDef(115, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(115, 9, xls.AddFormat(fmt));
            xls.SetCellValue(115, 9, new TFormula("=C115*H115"));

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
            xls.SetCellValue(116, 3, 0.21);
            xls.SetCellValue(116, 5, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(116, 6, 1);
            xls.SetCellValue(116, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(116, 8);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(116, 8, xls.AddFormat(fmt));
            xls.SetCellValue(116, 8, 0.705019741);

            fmt = xls.GetCellVisibleFormatDef(116, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(116, 9, xls.AddFormat(fmt));
            xls.SetCellValue(116, 9, new TFormula("=C116*H116"));

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
            xls.SetCellValue(117, 3, 0);
            xls.SetCellValue(117, 5, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(117, 6, 1);
            xls.SetCellValue(117, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(117, 8);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(117, 8, xls.AddFormat(fmt));
            xls.SetCellValue(117, 8, 0.705019741);

            fmt = xls.GetCellVisibleFormatDef(117, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(117, 9, xls.AddFormat(fmt));
            xls.SetCellValue(117, 9, new TFormula("=C117*H117"));

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
            xls.SetCellValue(118, 3, 0.7);
            xls.SetCellValue(118, 5, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(118, 6, 1);
            xls.SetCellValue(118, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(118, 8);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(118, 8, xls.AddFormat(fmt));
            xls.SetCellValue(118, 8, 0.705019741);

            fmt = xls.GetCellVisibleFormatDef(118, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(118, 9, xls.AddFormat(fmt));
            xls.SetCellValue(118, 9, new TFormula("=C118*H118"));

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
            xls.SetCellValue(119, 3, 0);
            xls.SetCellValue(119, 5, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(119, 6, 1);
            xls.SetCellValue(119, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(119, 8);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(119, 8, xls.AddFormat(fmt));
            xls.SetCellValue(119, 8, 0.705019741);

            fmt = xls.GetCellVisibleFormatDef(119, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(119, 9, xls.AddFormat(fmt));
            xls.SetCellValue(119, 9, new TFormula("=C119*H119"));

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

            fmt = xls.GetCellVisibleFormatDef(120, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(120, 9, xls.AddFormat(fmt));

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

            fmt = xls.GetCellVisibleFormatDef(121, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(121, 9, xls.AddFormat(fmt));

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
            xls.SetCellValue(122, 3, 0);
            xls.SetCellValue(122, 5, new TFormula("='Gral Conf. Summary'!$H$33"));
            xls.SetCellValue(122, 6, 1);
            xls.SetCellValue(122, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(122, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(122, 8, xls.AddFormat(fmt));
            xls.SetCellValue(122, 8, new TFormula("=(  1 /  IF(E122<>1,VLOOKUP(E122,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)  ) "
            + "    *       IF(F122<>1,VLOOKUP(F122,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) *"
            + " IF(G122<>1,VLOOKUP(G122,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(122, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(122, 9, xls.AddFormat(fmt));
            xls.SetCellValue(122, 9, new TFormula("=C122*H122"));

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
            xls.SetCellValue(123, 3, 8255);
            xls.SetCellValue(123, 5, new TFormula("='Gral Conf. Summary'!$H$33"));
            xls.SetCellValue(123, 6, 1);
            xls.SetCellValue(123, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(123, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(123, 8, xls.AddFormat(fmt));
            xls.SetCellValue(123, 8, new TFormula("=(  1 /  IF(E123<>1,VLOOKUP(E123,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)  ) "
            + "    *       IF(F123<>1,VLOOKUP(F123,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) *"
            + " IF(G123<>1,VLOOKUP(G123,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(123, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(123, 9, xls.AddFormat(fmt));
            xls.SetCellValue(123, 9, new TFormula("=C123*H123"));

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
            xls.SetCellValue(124, 3, 2.7);
            xls.SetCellValue(124, 5, 1);
            xls.SetCellValue(124, 6, 1);
            xls.SetCellValue(124, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(124, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(124, 8, xls.AddFormat(fmt));
            xls.SetCellValue(124, 8, new TFormula("=(    1 / IF(E124<>1,VLOOKUP(E124,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)   )"
            + "      *       IF(F124<>1,VLOOKUP(F124,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)"
            + " * IF(G124<>1,VLOOKUP(G124,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(124, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(124, 9, xls.AddFormat(fmt));
            xls.SetCellValue(124, 9, new TFormula("=C124*H124"));

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

            fmt = xls.GetCellVisibleFormatDef(125, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(125, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(126, 2);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(126, 2, xls.AddFormat(fmt));
            xls.SetCellValue(126, 2, new TFormula("=+\"If you receive any credit from the cooperative to invest in your farm or coffee"
            + " production activities,  What was the amount of this credit in \"&'Gral Conf. Summary'!$H$33&\""
            + " ?\""));

            fmt = xls.GetCellVisibleFormatDef(126, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(126, 3, xls.AddFormat(fmt));
            xls.SetCellValue(126, 3, 14000);
            xls.SetCellValue(126, 5, new TFormula("='Gral Conf. Summary'!$H$33"));
            xls.SetCellValue(126, 6, 1);
            xls.SetCellValue(126, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(126, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(126, 8, xls.AddFormat(fmt));
            xls.SetCellValue(126, 8, new TFormula("=(  1 /  IF(E126<>1,VLOOKUP(E126,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)  ) "
            + "    *       IF(F126<>1,VLOOKUP(F126,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) *"
            + " IF(G126<>1,VLOOKUP(G126,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(126, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(126, 9, xls.AddFormat(fmt));
            xls.SetCellValue(126, 9, new TFormula("=C126*H126"));

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
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(127, 3, xls.AddFormat(fmt));
            xls.SetCellValue(127, 3, 1.59999999999991);
            xls.SetCellValue(127, 5, 1);
            xls.SetCellValue(127, 6, 1);
            xls.SetCellValue(127, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(127, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(127, 8, xls.AddFormat(fmt));
            xls.SetCellValue(127, 8, new TFormula("=(    1 / IF(E127<>1,VLOOKUP(E127,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)   )"
            + "      *       IF(F127<>1,VLOOKUP(F127,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)"
            + " * IF(G127<>1,VLOOKUP(G127,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(127, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(127, 9, xls.AddFormat(fmt));
            xls.SetCellValue(127, 9, new TFormula("=C127*H127"));

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
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(128, 3, xls.AddFormat(fmt));
            xls.SetCellValue(128, 3, 1.01);
            xls.SetCellValue(128, 5, 1);
            xls.SetCellValue(128, 6, 1);
            xls.SetCellValue(128, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(128, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(128, 8, xls.AddFormat(fmt));
            xls.SetCellValue(128, 8, new TFormula("=(    1 / IF(E128<>1,VLOOKUP(E128,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)   )"
            + "      *       IF(F128<>1,VLOOKUP(F128,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)"
            + " * IF(G128<>1,VLOOKUP(G128,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(128, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(128, 9, xls.AddFormat(fmt));
            xls.SetCellValue(128, 9, new TFormula("=C128*H128"));

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
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(129, 3, xls.AddFormat(fmt));
            xls.SetCellValue(129, 3, 5260);
            xls.SetCellValue(129, 5, new TFormula("='Gral Conf. Summary'!$H$33"));
            xls.SetCellValue(129, 6, 1);
            xls.SetCellValue(129, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(129, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(129, 8, xls.AddFormat(fmt));
            xls.SetCellValue(129, 8, new TFormula("=(  1 /  IF(E129<>1,VLOOKUP(E129,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)  ) "
            + "    *       IF(F129<>1,VLOOKUP(F129,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) *"
            + " IF(G129<>1,VLOOKUP(G129,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(129, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(129, 9, xls.AddFormat(fmt));
            xls.SetCellValue(129, 9, new TFormula("=C129*H129"));

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
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(130, 3, xls.AddFormat(fmt));
            xls.SetCellValue(130, 3, 2.20000000000005);
            xls.SetCellValue(130, 5, 1);
            xls.SetCellValue(130, 6, 1);
            xls.SetCellValue(130, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(130, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(130, 8, xls.AddFormat(fmt));
            xls.SetCellValue(130, 8, new TFormula("=(    1 / IF(E130<>1,VLOOKUP(E130,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)   )"
            + "      *       IF(F130<>1,VLOOKUP(F130,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)"
            + " * IF(G130<>1,VLOOKUP(G130,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(130, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(130, 9, xls.AddFormat(fmt));
            xls.SetCellValue(130, 9, new TFormula("=C130*H130"));

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
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(131, 3, xls.AddFormat(fmt));
            xls.SetCellValue(131, 3, 3.21);
            xls.SetCellValue(131, 5, 1);
            xls.SetCellValue(131, 6, 1);
            xls.SetCellValue(131, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(131, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(131, 8, xls.AddFormat(fmt));
            xls.SetCellValue(131, 8, new TFormula("= IF(E131<>1,VLOOKUP(E131,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)       *   "
            + "    IF(F131<>1,VLOOKUP(F131,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) * IF(G131<>1,VLOOKUP(G131,'Gral"
            + " Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(131, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(131, 9, xls.AddFormat(fmt));
            xls.SetCellValue(131, 9, new TFormula("=C131*H131"));

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

            fmt = xls.GetCellVisibleFormatDef(132, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(132, 9, xls.AddFormat(fmt));
            xls.SetCellValue(132, 9, new TFormula("=C132*H132"));

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

            fmt = xls.GetCellVisibleFormatDef(133, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(133, 9, xls.AddFormat(fmt));
            xls.SetCellValue(133, 9, new TFormula("=C133*H133"));

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

            fmt = xls.GetCellVisibleFormatDef(134, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(134, 9, xls.AddFormat(fmt));
            xls.SetCellValue(134, 9, new TFormula("=C134*H134"));

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
            xls.SetCellValue(135, 3, 487);
            xls.SetCellValue(135, 5, new TFormula("='Gral Conf. Summary'!$H$33"));
            xls.SetCellValue(135, 6, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(135, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(135, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(135, 8, xls.AddFormat(fmt));
            xls.SetCellValue(135, 8, new TFormula("=(    1/ IF(E135<>1,VLOOKUP(E135,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) )  "
            + "    *      IF(F135<>1,VLOOKUP(F135,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)  "
            + "    * IF(G135<>1,VLOOKUP(G135,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(135, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(135, 9, xls.AddFormat(fmt));
            xls.SetCellValue(135, 9, new TFormula("=C135*H135"));

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
            xls.SetCellValue(136, 3, 430);
            xls.SetCellValue(136, 5, new TFormula("='Gral Conf. Summary'!$H$33"));
            xls.SetCellValue(136, 6, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(136, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(136, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(136, 8, xls.AddFormat(fmt));
            xls.SetCellValue(136, 8, new TFormula("=(    1/ IF(E136<>1,VLOOKUP(E136,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) )  "
            + "    *      IF(F136<>1,VLOOKUP(F136,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)  "
            + "    * IF(G136<>1,VLOOKUP(G136,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(136, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(136, 9, xls.AddFormat(fmt));
            xls.SetCellValue(136, 9, new TFormula("=C136*H136"));

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
            xls.SetCellValue(137, 3, 630);
            xls.SetCellValue(137, 5, new TFormula("='Gral Conf. Summary'!$H$33"));
            xls.SetCellValue(137, 6, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(137, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(137, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(137, 8, xls.AddFormat(fmt));
            xls.SetCellValue(137, 8, new TFormula("=(    1/ IF(E137<>1,VLOOKUP(E137,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) )  "
            + "    *      IF(F137<>1,VLOOKUP(F137,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)  "
            + "    * IF(G137<>1,VLOOKUP(G137,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(137, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(137, 9, xls.AddFormat(fmt));
            xls.SetCellValue(137, 9, new TFormula("=C137*H137"));

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
            xls.SetCellValue(138, 3, 0);
            xls.SetCellValue(138, 5, new TFormula("='Gral Conf. Summary'!$H$33"));
            xls.SetCellValue(138, 6, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(138, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(138, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(138, 8, xls.AddFormat(fmt));
            xls.SetCellValue(138, 8, new TFormula("=(    1/ IF(E138<>1,VLOOKUP(E138,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) )  "
            + "    *      IF(F138<>1,VLOOKUP(F138,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)  "
            + "    * IF(G138<>1,VLOOKUP(G138,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(138, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(138, 9, xls.AddFormat(fmt));
            xls.SetCellValue(138, 9, new TFormula("=C138*H138"));

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
            xls.SetCellValue(139, 3, 70);
            xls.SetCellValue(139, 5, new TFormula("='Gral Conf. Summary'!$H$33"));
            xls.SetCellValue(139, 6, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(139, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(139, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(139, 8, xls.AddFormat(fmt));
            xls.SetCellValue(139, 8, new TFormula("=(    1/ IF(E139<>1,VLOOKUP(E139,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) )  "
            + "    *      IF(F139<>1,VLOOKUP(F139,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)  "
            + "    * IF(G139<>1,VLOOKUP(G139,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(139, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(139, 9, xls.AddFormat(fmt));
            xls.SetCellValue(139, 9, new TFormula("=C139*H139"));

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
            xls.SetCellValue(140, 3, 80);
            xls.SetCellValue(140, 5, new TFormula("='Gral Conf. Summary'!$H$33"));
            xls.SetCellValue(140, 6, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(140, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(140, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(140, 8, xls.AddFormat(fmt));
            xls.SetCellValue(140, 8, new TFormula("=(    1/ IF(E140<>1,VLOOKUP(E140,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) )  "
            + "    *      IF(F140<>1,VLOOKUP(F140,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)  "
            + "    * IF(G140<>1,VLOOKUP(G140,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(140, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(140, 9, xls.AddFormat(fmt));
            xls.SetCellValue(140, 9, new TFormula("=C140*H140"));

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
            xls.SetCellValue(141, 3, 1510);
            xls.SetCellValue(141, 5, new TFormula("='Gral Conf. Summary'!$H$33"));
            xls.SetCellValue(141, 6, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(141, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(141, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(141, 8, xls.AddFormat(fmt));
            xls.SetCellValue(141, 8, new TFormula("=(    1/ IF(E141<>1,VLOOKUP(E141,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) )  "
            + "    *      IF(F141<>1,VLOOKUP(F141,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)  "
            + "    * IF(G141<>1,VLOOKUP(G141,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(141, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(141, 9, xls.AddFormat(fmt));
            xls.SetCellValue(141, 9, new TFormula("=C141*H141"));

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
            xls.SetCellValue(142, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(142, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(142, 9, xls.AddFormat(fmt));
            xls.SetCellValue(142, 9, new TFormula("=C142*H142"));

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
            xls.SetCellValue(143, 3, 2228);
            xls.SetCellValue(143, 5, new TFormula("='Gral Conf. Summary'!$H$33"));
            xls.SetCellValue(143, 6, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(143, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(143, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(143, 8, xls.AddFormat(fmt));
            xls.SetCellValue(143, 8, new TFormula("=(    1/ IF(E143<>1,VLOOKUP(E143,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) )  "
            + "    *      IF(F143<>1,VLOOKUP(F143,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)  "
            + "    * IF(G143<>1,VLOOKUP(G143,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(143, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(143, 9, xls.AddFormat(fmt));
            xls.SetCellValue(143, 9, new TFormula("=C143*H143"));

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
            xls.SetCellValue(144, 3, 979.7);
            xls.SetCellValue(144, 5, new TFormula("='Gral Conf. Summary'!$H$33"));
            xls.SetCellValue(144, 6, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(144, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(144, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(144, 8, xls.AddFormat(fmt));
            xls.SetCellValue(144, 8, new TFormula("=(    1/ IF(E144<>1,VLOOKUP(E144,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) )  "
            + "    *      IF(F144<>1,VLOOKUP(F144,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)  "
            + "    * IF(G144<>1,VLOOKUP(G144,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(144, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(144, 9, xls.AddFormat(fmt));
            xls.SetCellValue(144, 9, new TFormula("=C144*H144"));

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
            xls.SetCellValue(145, 3, 1815);
            xls.SetCellValue(145, 5, new TFormula("='Gral Conf. Summary'!$H$33"));
            xls.SetCellValue(145, 6, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(145, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(145, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(145, 8, xls.AddFormat(fmt));
            xls.SetCellValue(145, 8, new TFormula("=(    1/ IF(E145<>1,VLOOKUP(E145,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) )  "
            + "    *      IF(F145<>1,VLOOKUP(F145,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)  "
            + "    * IF(G145<>1,VLOOKUP(G145,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(145, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(145, 9, xls.AddFormat(fmt));
            xls.SetCellValue(145, 9, new TFormula("=C145*H145"));

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
            xls.SetCellValue(146, 3, 391);
            xls.SetCellValue(146, 5, new TFormula("='Gral Conf. Summary'!$H$33"));
            xls.SetCellValue(146, 6, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(146, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(146, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(146, 8, xls.AddFormat(fmt));
            xls.SetCellValue(146, 8, new TFormula("=(    1/ IF(E146<>1,VLOOKUP(E146,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) )  "
            + "    *      IF(F146<>1,VLOOKUP(F146,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)  "
            + "    * IF(G146<>1,VLOOKUP(G146,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(146, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(146, 9, xls.AddFormat(fmt));
            xls.SetCellValue(146, 9, new TFormula("=C146*H146"));

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
            xls.SetCellValue(147, 3, 240);
            xls.SetCellValue(147, 5, new TFormula("='Gral Conf. Summary'!$H$33"));
            xls.SetCellValue(147, 6, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(147, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(147, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(147, 8, xls.AddFormat(fmt));
            xls.SetCellValue(147, 8, new TFormula("=(    1/ IF(E147<>1,VLOOKUP(E147,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) )  "
            + "    *      IF(F147<>1,VLOOKUP(F147,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)  "
            + "    * IF(G147<>1,VLOOKUP(G147,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(147, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(147, 9, xls.AddFormat(fmt));
            xls.SetCellValue(147, 9, new TFormula("=C147*H147"));

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
            xls.SetCellValue(148, 3, 1066);
            xls.SetCellValue(148, 5, new TFormula("='Gral Conf. Summary'!$H$33"));
            xls.SetCellValue(148, 6, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(148, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(148, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(148, 8, xls.AddFormat(fmt));
            xls.SetCellValue(148, 8, new TFormula("=(    1/ IF(E148<>1,VLOOKUP(E148,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) )  "
            + "    *      IF(F148<>1,VLOOKUP(F148,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)  "
            + "    * IF(G148<>1,VLOOKUP(G148,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(148, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(148, 9, xls.AddFormat(fmt));
            xls.SetCellValue(148, 9, new TFormula("=C148*H148"));

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
            xls.SetCellValue(149, 3, 38.25);
            xls.SetCellValue(149, 5, new TFormula("='Gral Conf. Summary'!$H$33"));
            xls.SetCellValue(149, 6, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(149, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(149, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(149, 8, xls.AddFormat(fmt));
            xls.SetCellValue(149, 8, new TFormula("=(    1/ IF(E149<>1,VLOOKUP(E149,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) )  "
            + "    *      IF(F149<>1,VLOOKUP(F149,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)  "
            + "    * IF(G149<>1,VLOOKUP(G149,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(149, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(149, 9, xls.AddFormat(fmt));
            xls.SetCellValue(149, 9, new TFormula("=C149*H149"));

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
            xls.SetCellValue(150, 3, 3436);
            xls.SetCellValue(150, 5, new TFormula("='Gral Conf. Summary'!$H$33"));
            xls.SetCellValue(150, 6, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(150, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(150, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(150, 8, xls.AddFormat(fmt));
            xls.SetCellValue(150, 8, new TFormula("=(    1/ IF(E150<>1,VLOOKUP(E150,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) )  "
            + "    *      IF(F150<>1,VLOOKUP(F150,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)  "
            + "    * IF(G150<>1,VLOOKUP(G150,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(150, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(150, 9, xls.AddFormat(fmt));
            xls.SetCellValue(150, 9, new TFormula("=C150*H150"));

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
            xls.SetCellValue(151, 3, 482.5);
            xls.SetCellValue(151, 5, new TFormula("='Gral Conf. Summary'!$H$33"));
            xls.SetCellValue(151, 6, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(151, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(151, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(151, 8, xls.AddFormat(fmt));
            xls.SetCellValue(151, 8, new TFormula("=(    1/ IF(E151<>1,VLOOKUP(E151,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) )  "
            + "    *      IF(F151<>1,VLOOKUP(F151,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)  "
            + "    * IF(G151<>1,VLOOKUP(G151,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(151, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(151, 9, xls.AddFormat(fmt));
            xls.SetCellValue(151, 9, new TFormula("=C151*H151"));

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
            xls.SetCellValue(152, 3, 0);
            xls.SetCellValue(152, 5, new TFormula("='Gral Conf. Summary'!$H$33"));
            xls.SetCellValue(152, 6, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(152, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(152, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(152, 8, xls.AddFormat(fmt));
            xls.SetCellValue(152, 8, new TFormula("=(    1/ IF(E152<>1,VLOOKUP(E152,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) )  "
            + "    *      IF(F152<>1,VLOOKUP(F152,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)  "
            + "    * IF(G152<>1,VLOOKUP(G152,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(152, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(152, 9, xls.AddFormat(fmt));
            xls.SetCellValue(152, 9, new TFormula("=C152*H152"));

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
            xls.SetCellValue(153, 3, 240);
            xls.SetCellValue(153, 5, new TFormula("='Gral Conf. Summary'!$H$33"));
            xls.SetCellValue(153, 6, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(153, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(153, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(153, 8, xls.AddFormat(fmt));
            xls.SetCellValue(153, 8, new TFormula("=(    1/ IF(E153<>1,VLOOKUP(E153,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) )  "
            + "    *      IF(F153<>1,VLOOKUP(F153,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)  "
            + "    * IF(G153<>1,VLOOKUP(G153,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(153, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(153, 9, xls.AddFormat(fmt));
            xls.SetCellValue(153, 9, new TFormula("=C153*H153"));

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
            xls.SetCellValue(154, 2, "Forsforic rock");

            fmt = xls.GetCellVisibleFormatDef(154, 3);
            fmt.Font.Size20 = 280;
            fmt.Format = "0.00";
            xls.SetCellFormat(154, 3, xls.AddFormat(fmt));
            xls.SetCellValue(154, 3, 0);
            xls.SetCellValue(154, 5, new TFormula("='Gral Conf. Summary'!$H$33"));
            xls.SetCellValue(154, 6, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(154, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(154, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(154, 8, xls.AddFormat(fmt));
            xls.SetCellValue(154, 8, new TFormula("=(    1/ IF(E154<>1,VLOOKUP(E154,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) )  "
            + "    *      IF(F154<>1,VLOOKUP(F154,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)  "
            + "    * IF(G154<>1,VLOOKUP(G154,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(154, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(154, 9, xls.AddFormat(fmt));
            xls.SetCellValue(154, 9, new TFormula("=C154*H154"));

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
            xls.SetCellValue(155, 3, 575.5);
            xls.SetCellValue(155, 5, new TFormula("='Gral Conf. Summary'!$H$33"));
            xls.SetCellValue(155, 6, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(155, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(155, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(155, 8, xls.AddFormat(fmt));
            xls.SetCellValue(155, 8, new TFormula("=(    1/ IF(E155<>1,VLOOKUP(E155,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) )  "
            + "    *      IF(F155<>1,VLOOKUP(F155,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)  "
            + "    * IF(G155<>1,VLOOKUP(G155,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(155, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(155, 9, xls.AddFormat(fmt));
            xls.SetCellValue(155, 9, new TFormula("=C155*H155"));

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

            fmt = xls.GetCellVisibleFormatDef(156, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(156, 9, xls.AddFormat(fmt));

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
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xCC, 0xFF, 0xCC);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(157, 3, xls.AddFormat(fmt));
            xls.SetCellValue(157, 3, 3517.98883137063);
            xls.SetCellValue(157, 5, new TFormula("='Gral Conf. Summary'!$H$33"));
            xls.SetCellValue(157, 6, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(157, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(157, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(157, 8, xls.AddFormat(fmt));
            xls.SetCellValue(157, 8, new TFormula("=(    1/ IF(E157<>1,VLOOKUP(E157,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) )  "
            + "    *      IF(F157<>1,VLOOKUP(F157,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)  "
            + "    * IF(G157<>1,VLOOKUP(G157,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(157, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(157, 9, xls.AddFormat(fmt));
            xls.SetCellValue(157, 9, new TFormula("=C157*H157"));

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
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xCC, 0xFF, 0xCC);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(158, 3, xls.AddFormat(fmt));
            xls.SetCellValue(158, 3, 3517.98883137063);
            xls.SetCellValue(158, 5, new TFormula("='Gral Conf. Summary'!$H$33"));
            xls.SetCellValue(158, 6, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(158, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(158, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(158, 8, xls.AddFormat(fmt));
            xls.SetCellValue(158, 8, new TFormula("=(    1/ IF(E158<>1,VLOOKUP(E158,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) )  "
            + "    *      IF(F158<>1,VLOOKUP(F158,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)  "
            + "    * IF(G158<>1,VLOOKUP(G158,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(158, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(158, 9, xls.AddFormat(fmt));
            xls.SetCellValue(158, 9, new TFormula("=C158*H158"));

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

            fmt = xls.GetCellVisibleFormatDef(159, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(159, 9, xls.AddFormat(fmt));

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
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xCC, 0xFF, 0xCC);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(160, 3, xls.AddFormat(fmt));
            xls.SetCellValue(160, 3, 1037.79389709306);
            xls.SetCellValue(160, 5, new TFormula("='Gral Conf. Summary'!$H$33"));
            xls.SetCellValue(160, 6, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(160, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(160, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(160, 8, xls.AddFormat(fmt));
            xls.SetCellValue(160, 8, new TFormula("=(    1/ IF(E160<>1,VLOOKUP(E160,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) )  "
            + "    *      IF(F160<>1,VLOOKUP(F160,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)  "
            + "    * IF(G160<>1,VLOOKUP(G160,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(160, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(160, 9, xls.AddFormat(fmt));
            xls.SetCellValue(160, 9, new TFormula("=C160*H160"));

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
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xCC, 0xFF, 0xCC);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(161, 3, xls.AddFormat(fmt));
            xls.SetCellValue(161, 3, 1037.79389709306);
            xls.SetCellValue(161, 5, new TFormula("='Gral Conf. Summary'!$H$33"));
            xls.SetCellValue(161, 6, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(161, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(161, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(161, 8, xls.AddFormat(fmt));
            xls.SetCellValue(161, 8, new TFormula("=(    1/ IF(E161<>1,VLOOKUP(E161,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) )  "
            + "    *      IF(F161<>1,VLOOKUP(F161,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)  "
            + "    * IF(G161<>1,VLOOKUP(G161,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(161, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(161, 9, xls.AddFormat(fmt));
            xls.SetCellValue(161, 9, new TFormula("=C161*H161"));

            fmt = xls.GetCellVisibleFormatDef(162, 2);
            fmt.Font.Size20 = 220;
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Family = 0;
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
            fmt.Font.Size20 = 220;
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Family = 0;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(162, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(162, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(162, 9, xls.AddFormat(fmt));

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
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xCC, 0xFF, 0xCC);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(163, 3, xls.AddFormat(fmt));
            xls.SetCellValue(163, 3, 0);
            xls.SetCellValue(163, 5, new TFormula("='Gral Conf. Summary'!$H$33"));
            xls.SetCellValue(163, 6, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(163, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(163, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(163, 8, xls.AddFormat(fmt));
            xls.SetCellValue(163, 8, new TFormula("=(    1/ IF(E163<>1,VLOOKUP(E163,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) )  "
            + "    *      IF(F163<>1,VLOOKUP(F163,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)  "
            + "    * IF(G163<>1,VLOOKUP(G163,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(163, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(163, 9, xls.AddFormat(fmt));
            xls.SetCellValue(163, 9, 0);

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
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xCC, 0xFF, 0xCC);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(164, 3, xls.AddFormat(fmt));
            xls.SetCellValue(164, 3, 0);
            xls.SetCellValue(164, 5, new TFormula("='Gral Conf. Summary'!$H$33"));
            xls.SetCellValue(164, 6, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(164, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(164, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(164, 8, xls.AddFormat(fmt));
            xls.SetCellValue(164, 8, new TFormula("=(    1/ IF(E164<>1,VLOOKUP(E164,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) )  "
            + "    *      IF(F164<>1,VLOOKUP(F164,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)  "
            + "    * IF(G164<>1,VLOOKUP(G164,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(164, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(164, 9, xls.AddFormat(fmt));
            xls.SetCellValue(164, 9, new TFormula("=C164*H164"));

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
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xCC, 0xFF, 0xCC);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(165, 3, xls.AddFormat(fmt));
            xls.SetCellValue(165, 3, 0);
            xls.SetCellValue(165, 5, new TFormula("='Gral Conf. Summary'!$H$33"));
            xls.SetCellValue(165, 6, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(165, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(165, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(165, 8, xls.AddFormat(fmt));
            xls.SetCellValue(165, 8, new TFormula("=(    1/ IF(E165<>1,VLOOKUP(E165,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) )  "
            + "    *      IF(F165<>1,VLOOKUP(F165,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)  "
            + "    * IF(G165<>1,VLOOKUP(G165,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(165, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(165, 9, xls.AddFormat(fmt));
            xls.SetCellValue(165, 9, new TFormula("=C165*H165"));

            fmt = xls.GetCellVisibleFormatDef(166, 2);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(166, 2, xls.AddFormat(fmt));
            xls.SetCellValue(166, 2, "Gas / fuel");
            xls.SetCellValue(166, 3, 0);
            xls.SetCellValue(166, 5, new TFormula("='Gral Conf. Summary'!$H$33"));
            xls.SetCellValue(166, 6, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(166, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(166, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(166, 8, xls.AddFormat(fmt));
            xls.SetCellValue(166, 8, new TFormula("=(    1/ IF(E166<>1,VLOOKUP(E166,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) )  "
            + "    *      IF(F166<>1,VLOOKUP(F166,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)  "
            + "    * IF(G166<>1,VLOOKUP(G166,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(166, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(166, 9, xls.AddFormat(fmt));
            xls.SetCellValue(166, 9, new TFormula("=C166*H166"));

            fmt = xls.GetCellVisibleFormatDef(167, 2);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(167, 2, xls.AddFormat(fmt));
            xls.SetCellValue(167, 2, "Other inputs for mantainance");
            xls.SetCellValue(167, 3, 0);
            xls.SetCellValue(167, 5, new TFormula("='Gral Conf. Summary'!$H$33"));
            xls.SetCellValue(167, 6, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(167, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(167, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(167, 8, xls.AddFormat(fmt));
            xls.SetCellValue(167, 8, new TFormula("=(    1/ IF(E167<>1,VLOOKUP(E167,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) )  "
            + "    *      IF(F167<>1,VLOOKUP(F167,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)  "
            + "    * IF(G167<>1,VLOOKUP(G167,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(167, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(167, 9, xls.AddFormat(fmt));
            xls.SetCellValue(167, 9, new TFormula("=C167*H167"));

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

            fmt = xls.GetCellVisibleFormatDef(168, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(168, 9, xls.AddFormat(fmt));
            xls.SetCellValue(168, 9, new TFormula("=C168*H168"));

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

            fmt = xls.GetCellVisibleFormatDef(169, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(169, 9, xls.AddFormat(fmt));

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

            fmt = xls.GetCellVisibleFormatDef(170, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(170, 9, xls.AddFormat(fmt));
            xls.SetCellValue(170, 9, new TFormula("=C170*H170"));

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
            xls.SetCellValue(171, 3, 1434);
            xls.SetCellValue(171, 5, new TFormula("='Gral Conf. Summary'!$H$33"));
            xls.SetCellValue(171, 6, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(171, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(171, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(171, 8, xls.AddFormat(fmt));
            xls.SetCellValue(171, 8, new TFormula("=(    1/ IF(E171<>1,VLOOKUP(E171,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) )  "
            + "    *      IF(F171<>1,VLOOKUP(F171,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)  "
            + "    * IF(G171<>1,VLOOKUP(G171,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(171, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(171, 9, xls.AddFormat(fmt));
            xls.SetCellValue(171, 9, new TFormula("=C171*H171"));

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
            xls.SetCellValue(172, 3, 5.36);
            xls.SetCellValue(172, 5, 1);
            xls.SetCellValue(172, 6, 1);
            xls.SetCellValue(172, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(172, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(172, 8, xls.AddFormat(fmt));
            xls.SetCellValue(172, 8, new TFormula("=(    1/ IF(E172<>1,VLOOKUP(E172,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) )  "
            + "    *      IF(F172<>1,VLOOKUP(F172,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)  "
            + "    * IF(G172<>1,VLOOKUP(G172,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(172, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(172, 9, xls.AddFormat(fmt));
            xls.SetCellValue(172, 9, new TFormula("=C172*H172"));

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
            xls.SetCellValue(173, 3, 340);
            xls.SetCellValue(173, 5, new TFormula("='Gral Conf. Summary'!$H$33"));
            xls.SetCellValue(173, 6, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(173, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(173, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(173, 8, xls.AddFormat(fmt));
            xls.SetCellValue(173, 8, new TFormula("=(    1/ IF(E173<>1,VLOOKUP(E173,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) )  "
            + "    *      IF(F173<>1,VLOOKUP(F173,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)  "
            + "    * IF(G173<>1,VLOOKUP(G173,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(173, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(173, 9, xls.AddFormat(fmt));
            xls.SetCellValue(173, 9, new TFormula("=C173*H173"));

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
            xls.SetCellValue(174, 3, 1.29);
            xls.SetCellValue(174, 5, 1);
            xls.SetCellValue(174, 6, 1);
            xls.SetCellValue(174, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(174, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(174, 8, xls.AddFormat(fmt));
            xls.SetCellValue(174, 8, new TFormula("=(    1/ IF(E174<>1,VLOOKUP(E174,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) )  "
            + "    *      IF(F174<>1,VLOOKUP(F174,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)  "
            + "    * IF(G174<>1,VLOOKUP(G174,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(174, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(174, 9, xls.AddFormat(fmt));
            xls.SetCellValue(174, 9, new TFormula("=C174*H174"));

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
            xls.SetCellValue(175, 3, 184);
            xls.SetCellValue(175, 5, new TFormula("='Gral Conf. Summary'!$H$33"));
            xls.SetCellValue(175, 6, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(175, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(175, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(175, 8, xls.AddFormat(fmt));
            xls.SetCellValue(175, 8, new TFormula("=(    1/ IF(E175<>1,VLOOKUP(E175,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) )  "
            + "    *      IF(F175<>1,VLOOKUP(F175,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)  "
            + "    * IF(G175<>1,VLOOKUP(G175,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(175, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(175, 9, xls.AddFormat(fmt));
            xls.SetCellValue(175, 9, new TFormula("=C175*H175"));

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
            xls.SetCellValue(176, 3, 4.09);
            xls.SetCellValue(176, 5, 1);
            xls.SetCellValue(176, 6, 1);
            xls.SetCellValue(176, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(176, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(176, 8, xls.AddFormat(fmt));
            xls.SetCellValue(176, 8, new TFormula("=(    1/ IF(E176<>1,VLOOKUP(E176,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) )  "
            + "    *      IF(F176<>1,VLOOKUP(F176,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)  "
            + "    * IF(G176<>1,VLOOKUP(G176,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(176, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(176, 9, xls.AddFormat(fmt));
            xls.SetCellValue(176, 9, new TFormula("=C176*H176"));

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
            xls.SetCellValue(177, 3, 190);
            xls.SetCellValue(177, 5, new TFormula("='Gral Conf. Summary'!$H$33"));
            xls.SetCellValue(177, 6, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(177, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(177, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(177, 8, xls.AddFormat(fmt));
            xls.SetCellValue(177, 8, new TFormula("=(    1/ IF(E177<>1,VLOOKUP(E177,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) )  "
            + "    *      IF(F177<>1,VLOOKUP(F177,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)  "
            + "    * IF(G177<>1,VLOOKUP(G177,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(177, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(177, 9, xls.AddFormat(fmt));
            xls.SetCellValue(177, 9, new TFormula("=C177*H177"));

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
            xls.SetCellValue(178, 3, 4.8);
            xls.SetCellValue(178, 5, 1);
            xls.SetCellValue(178, 6, 1);
            xls.SetCellValue(178, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(178, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(178, 8, xls.AddFormat(fmt));
            xls.SetCellValue(178, 8, new TFormula("=(    1/ IF(E178<>1,VLOOKUP(E178,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) )  "
            + "    *      IF(F178<>1,VLOOKUP(F178,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)  "
            + "    * IF(G178<>1,VLOOKUP(G178,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(178, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(178, 9, xls.AddFormat(fmt));
            xls.SetCellValue(178, 9, new TFormula("=C178*H178"));

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
            xls.SetCellValue(179, 3, 943);
            xls.SetCellValue(179, 5, new TFormula("='Gral Conf. Summary'!$H$33"));
            xls.SetCellValue(179, 6, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(179, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(179, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(179, 8, xls.AddFormat(fmt));
            xls.SetCellValue(179, 8, new TFormula("=(    1/ IF(E179<>1,VLOOKUP(E179,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) )  "
            + "    *      IF(F179<>1,VLOOKUP(F179,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)  "
            + "    * IF(G179<>1,VLOOKUP(G179,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(179, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(179, 9, xls.AddFormat(fmt));
            xls.SetCellValue(179, 9, new TFormula("=C179*H179"));

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
            xls.SetCellValue(180, 3, 4.84);
            xls.SetCellValue(180, 5, 1);
            xls.SetCellValue(180, 6, 1);
            xls.SetCellValue(180, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(180, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(180, 8, xls.AddFormat(fmt));
            xls.SetCellValue(180, 8, new TFormula("=(    1/ IF(E180<>1,VLOOKUP(E180,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) )  "
            + "    *      IF(F180<>1,VLOOKUP(F180,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)  "
            + "    * IF(G180<>1,VLOOKUP(G180,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(180, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(180, 9, xls.AddFormat(fmt));
            xls.SetCellValue(180, 9, new TFormula("=C180*H180"));

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
            xls.SetCellValue(181, 3, 289);
            xls.SetCellValue(181, 5, new TFormula("='Gral Conf. Summary'!$H$33"));
            xls.SetCellValue(181, 6, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(181, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(181, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(181, 8, xls.AddFormat(fmt));
            xls.SetCellValue(181, 8, new TFormula("=(    1/ IF(E181<>1,VLOOKUP(E181,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) )  "
            + "    *      IF(F181<>1,VLOOKUP(F181,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)  "
            + "    * IF(G181<>1,VLOOKUP(G181,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(181, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(181, 9, xls.AddFormat(fmt));
            xls.SetCellValue(181, 9, new TFormula("=C181*H181"));

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
            xls.SetCellValue(182, 3, 0.63);
            xls.SetCellValue(182, 5, 1);
            xls.SetCellValue(182, 6, 1);
            xls.SetCellValue(182, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(182, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(182, 8, xls.AddFormat(fmt));
            xls.SetCellValue(182, 8, new TFormula("=(    1/ IF(E182<>1,VLOOKUP(E182,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) )  "
            + "    *      IF(F182<>1,VLOOKUP(F182,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)  "
            + "    * IF(G182<>1,VLOOKUP(G182,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(182, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(182, 9, xls.AddFormat(fmt));
            xls.SetCellValue(182, 9, new TFormula("=C182*H182"));

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
            xls.SetCellValue(183, 3, 210);
            xls.SetCellValue(183, 5, new TFormula("='Gral Conf. Summary'!$H$33"));
            xls.SetCellValue(183, 6, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(183, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(183, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(183, 8, xls.AddFormat(fmt));
            xls.SetCellValue(183, 8, new TFormula("=(    1/ IF(E183<>1,VLOOKUP(E183,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) )  "
            + "    *      IF(F183<>1,VLOOKUP(F183,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)  "
            + "    * IF(G183<>1,VLOOKUP(G183,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(183, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(183, 9, xls.AddFormat(fmt));
            xls.SetCellValue(183, 9, new TFormula("=C183*H183"));

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
            xls.SetCellValue(184, 3, 4.15);
            xls.SetCellValue(184, 5, 1);
            xls.SetCellValue(184, 6, 1);
            xls.SetCellValue(184, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(184, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(184, 8, xls.AddFormat(fmt));
            xls.SetCellValue(184, 8, new TFormula("=(    1/ IF(E184<>1,VLOOKUP(E184,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) )  "
            + "    *      IF(F184<>1,VLOOKUP(F184,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)  "
            + "    * IF(G184<>1,VLOOKUP(G184,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(184, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(184, 9, xls.AddFormat(fmt));
            xls.SetCellValue(184, 9, new TFormula("=C184*H184"));

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
            xls.SetCellValue(185, 3, 282);
            xls.SetCellValue(185, 5, new TFormula("='Gral Conf. Summary'!$H$33"));
            xls.SetCellValue(185, 6, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(185, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(185, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(185, 8, xls.AddFormat(fmt));
            xls.SetCellValue(185, 8, new TFormula("=(    1/ IF(E185<>1,VLOOKUP(E185,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) )  "
            + "    *      IF(F185<>1,VLOOKUP(F185,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)  "
            + "    * IF(G185<>1,VLOOKUP(G185,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(185, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(185, 9, xls.AddFormat(fmt));
            xls.SetCellValue(185, 9, new TFormula("=C185*H185"));

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
            xls.SetCellValue(186, 3, 5.04);
            xls.SetCellValue(186, 5, 1);
            xls.SetCellValue(186, 6, 1);
            xls.SetCellValue(186, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(186, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(186, 8, xls.AddFormat(fmt));
            xls.SetCellValue(186, 8, new TFormula("=(    1/ IF(E186<>1,VLOOKUP(E186,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) )  "
            + "    *      IF(F186<>1,VLOOKUP(F186,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)  "
            + "    * IF(G186<>1,VLOOKUP(G186,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(186, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(186, 9, xls.AddFormat(fmt));
            xls.SetCellValue(186, 9, new TFormula("=C186*H186"));

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
            xls.SetCellValue(187, 3, 4409);
            xls.SetCellValue(187, 5, new TFormula("='Gral Conf. Summary'!$H$33"));
            xls.SetCellValue(187, 6, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(187, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(187, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(187, 8, xls.AddFormat(fmt));
            xls.SetCellValue(187, 8, new TFormula("=(    1/ IF(E187<>1,VLOOKUP(E187,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) )  "
            + "    *      IF(F187<>1,VLOOKUP(F187,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)  "
            + "    * IF(G187<>1,VLOOKUP(G187,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(187, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(187, 9, xls.AddFormat(fmt));
            xls.SetCellValue(187, 9, new TFormula("=C187*H187"));

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
            xls.SetCellValue(188, 3, 3.94);
            xls.SetCellValue(188, 5, 1);
            xls.SetCellValue(188, 6, 1);
            xls.SetCellValue(188, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(188, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(188, 8, xls.AddFormat(fmt));
            xls.SetCellValue(188, 8, new TFormula("=(    1/ IF(E188<>1,VLOOKUP(E188,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) )  "
            + "    *      IF(F188<>1,VLOOKUP(F188,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)  "
            + "    * IF(G188<>1,VLOOKUP(G188,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(188, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(188, 9, xls.AddFormat(fmt));
            xls.SetCellValue(188, 9, new TFormula("=C188*H188"));

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
            xls.SetCellValue(189, 3, 203);
            xls.SetCellValue(189, 5, new TFormula("='Gral Conf. Summary'!$H$33"));
            xls.SetCellValue(189, 6, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(189, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(189, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(189, 8, xls.AddFormat(fmt));
            xls.SetCellValue(189, 8, new TFormula("=(    1/ IF(E189<>1,VLOOKUP(E189,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) )  "
            + "    *      IF(F189<>1,VLOOKUP(F189,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)  "
            + "    * IF(G189<>1,VLOOKUP(G189,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(189, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(189, 9, xls.AddFormat(fmt));
            xls.SetCellValue(189, 9, new TFormula("=C189*H189"));

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
            xls.SetCellValue(190, 3, 1.6);
            xls.SetCellValue(190, 5, 1);
            xls.SetCellValue(190, 6, 1);
            xls.SetCellValue(190, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(190, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(190, 8, xls.AddFormat(fmt));
            xls.SetCellValue(190, 8, new TFormula("=(    1/ IF(E190<>1,VLOOKUP(E190,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) )  "
            + "    *      IF(F190<>1,VLOOKUP(F190,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)  "
            + "    * IF(G190<>1,VLOOKUP(G190,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(190, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(190, 9, xls.AddFormat(fmt));
            xls.SetCellValue(190, 9, new TFormula("=C190*H190"));

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
            xls.SetCellValue(191, 3, 8248);
            xls.SetCellValue(191, 5, new TFormula("='Gral Conf. Summary'!$H$33"));
            xls.SetCellValue(191, 6, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(191, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(191, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(191, 8, xls.AddFormat(fmt));
            xls.SetCellValue(191, 8, new TFormula("=(    1/ IF(E191<>1,VLOOKUP(E191,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) )  "
            + "    *      IF(F191<>1,VLOOKUP(F191,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)  "
            + "    * IF(G191<>1,VLOOKUP(G191,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(191, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(191, 9, xls.AddFormat(fmt));
            xls.SetCellValue(191, 9, new TFormula("=C191*H191"));

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
            xls.SetCellValue(192, 3, 7.03);
            xls.SetCellValue(192, 5, 1);
            xls.SetCellValue(192, 6, 1);
            xls.SetCellValue(192, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(192, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(192, 8, xls.AddFormat(fmt));
            xls.SetCellValue(192, 8, new TFormula("=(    1/ IF(E192<>1,VLOOKUP(E192,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) )  "
            + "    *      IF(F192<>1,VLOOKUP(F192,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)  "
            + "    * IF(G192<>1,VLOOKUP(G192,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(192, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(192, 9, xls.AddFormat(fmt));
            xls.SetCellValue(192, 9, new TFormula("=C192*H192"));

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
            xls.SetCellValue(193, 3, 190);
            xls.SetCellValue(193, 5, new TFormula("='Gral Conf. Summary'!$H$33"));
            xls.SetCellValue(193, 6, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(193, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(193, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(193, 8, xls.AddFormat(fmt));
            xls.SetCellValue(193, 8, new TFormula("=(    1/ IF(E193<>1,VLOOKUP(E193,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) )  "
            + "    *      IF(F193<>1,VLOOKUP(F193,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)  "
            + "    * IF(G193<>1,VLOOKUP(G193,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(193, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(193, 9, xls.AddFormat(fmt));
            xls.SetCellValue(193, 9, new TFormula("=C193*H193"));

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
            xls.SetCellValue(194, 3, 3.77);
            xls.SetCellValue(194, 5, 1);
            xls.SetCellValue(194, 6, 1);
            xls.SetCellValue(194, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(194, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(194, 8, xls.AddFormat(fmt));
            xls.SetCellValue(194, 8, new TFormula("=(    1/ IF(E194<>1,VLOOKUP(E194,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) )  "
            + "    *      IF(F194<>1,VLOOKUP(F194,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)  "
            + "    * IF(G194<>1,VLOOKUP(G194,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(194, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(194, 9, xls.AddFormat(fmt));
            xls.SetCellValue(194, 9, new TFormula("=C194*H194"));

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
            xls.SetCellValue(195, 3, 0);
            xls.SetCellValue(195, 5, new TFormula("='Gral Conf. Summary'!$H$33"));
            xls.SetCellValue(195, 6, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(195, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(195, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(195, 8, xls.AddFormat(fmt));
            xls.SetCellValue(195, 8, new TFormula("=(    1/ IF(E195<>1,VLOOKUP(E195,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) )  "
            + "    *      IF(F195<>1,VLOOKUP(F195,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)  "
            + "    * IF(G195<>1,VLOOKUP(G195,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(195, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(195, 9, xls.AddFormat(fmt));
            xls.SetCellValue(195, 9, new TFormula("=C195*H195"));

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
            xls.SetCellValue(196, 3, 0);
            xls.SetCellValue(196, 5, 1);
            xls.SetCellValue(196, 6, 1);
            xls.SetCellValue(196, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(196, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(196, 8, xls.AddFormat(fmt));
            xls.SetCellValue(196, 8, new TFormula("=(    1/ IF(E196<>1,VLOOKUP(E196,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) )  "
            + "    *      IF(F196<>1,VLOOKUP(F196,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)  "
            + "    * IF(G196<>1,VLOOKUP(G196,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(196, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(196, 9, xls.AddFormat(fmt));
            xls.SetCellValue(196, 9, new TFormula("=C196*H196"));

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
            xls.SetCellValue(197, 3, 267);
            xls.SetCellValue(197, 5, new TFormula("='Gral Conf. Summary'!$H$33"));
            xls.SetCellValue(197, 6, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(197, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(197, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(197, 8, xls.AddFormat(fmt));
            xls.SetCellValue(197, 8, new TFormula("=(    1/ IF(E197<>1,VLOOKUP(E197,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) )  "
            + "    *      IF(F197<>1,VLOOKUP(F197,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)  "
            + "    * IF(G197<>1,VLOOKUP(G197,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(197, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(197, 9, xls.AddFormat(fmt));
            xls.SetCellValue(197, 9, new TFormula("=C197*H197"));

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
            xls.SetCellValue(198, 3, 4.6);
            xls.SetCellValue(198, 5, 1);
            xls.SetCellValue(198, 6, 1);
            xls.SetCellValue(198, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(198, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(198, 8, xls.AddFormat(fmt));
            xls.SetCellValue(198, 8, new TFormula("=(    1/ IF(E198<>1,VLOOKUP(E198,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) )  "
            + "    *      IF(F198<>1,VLOOKUP(F198,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)  "
            + "    * IF(G198<>1,VLOOKUP(G198,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(198, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(198, 9, xls.AddFormat(fmt));
            xls.SetCellValue(198, 9, new TFormula("=C198*H198"));

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
            xls.SetCellValue(199, 3, 251);
            xls.SetCellValue(199, 5, new TFormula("='Gral Conf. Summary'!$H$33"));
            xls.SetCellValue(199, 6, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(199, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(199, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(199, 8, xls.AddFormat(fmt));
            xls.SetCellValue(199, 8, new TFormula("=(    1/ IF(E199<>1,VLOOKUP(E199,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) )  "
            + "    *      IF(F199<>1,VLOOKUP(F199,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)  "
            + "    * IF(G199<>1,VLOOKUP(G199,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(199, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(199, 9, xls.AddFormat(fmt));
            xls.SetCellValue(199, 9, new TFormula("=C199*H199"));

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
            xls.SetCellValue(200, 3, 7.65);
            xls.SetCellValue(200, 5, 1);
            xls.SetCellValue(200, 6, 1);
            xls.SetCellValue(200, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(200, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(200, 8, xls.AddFormat(fmt));
            xls.SetCellValue(200, 8, new TFormula("=(    1/ IF(E200<>1,VLOOKUP(E200,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) )  "
            + "    *      IF(F200<>1,VLOOKUP(F200,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)  "
            + "    * IF(G200<>1,VLOOKUP(G200,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(200, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(200, 9, xls.AddFormat(fmt));
            xls.SetCellValue(200, 9, new TFormula("=C200*H200"));

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

            fmt = xls.GetCellVisibleFormatDef(201, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(201, 9, xls.AddFormat(fmt));

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
            xls.SetCellValue(202, 3, 0);
            xls.SetCellValue(202, 5, new TFormula("='Gral Conf. Summary'!$H$33"));
            xls.SetCellValue(202, 6, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(202, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(202, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(202, 8, xls.AddFormat(fmt));
            xls.SetCellValue(202, 8, new TFormula("=(    1/ IF(E202<>1,VLOOKUP(E202,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) )  "
            + "    *      IF(F202<>1,VLOOKUP(F202,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)  "
            + "    * IF(G202<>1,VLOOKUP(G202,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(202, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(202, 9, xls.AddFormat(fmt));
            xls.SetCellValue(202, 9, new TFormula("=C202*H202"));

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
            xls.SetCellValue(203, 3, 8.14);
            xls.SetCellValue(203, 5, 1);
            xls.SetCellValue(203, 6, 1);
            xls.SetCellValue(203, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(203, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(203, 8, xls.AddFormat(fmt));
            xls.SetCellValue(203, 8, new TFormula("=(    1/ IF(E203<>1,VLOOKUP(E203,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) )  "
            + "    *      IF(F203<>1,VLOOKUP(F203,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)  "
            + "    * IF(G203<>1,VLOOKUP(G203,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(203, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(203, 9, xls.AddFormat(fmt));
            xls.SetCellValue(203, 9, new TFormula("=C203*H203"));

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
            xls.SetCellValue(204, 3, 78478);
            xls.SetCellValue(204, 5, new TFormula("='Gral Conf. Summary'!$H$33"));
            xls.SetCellValue(204, 6, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(204, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(204, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(204, 8, xls.AddFormat(fmt));
            xls.SetCellValue(204, 8, new TFormula("=(    1/ IF(E204<>1,VLOOKUP(E204,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) )  "
            + "    *      IF(F204<>1,VLOOKUP(F204,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)  "
            + "    * IF(G204<>1,VLOOKUP(G204,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(204, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(204, 9, xls.AddFormat(fmt));
            xls.SetCellValue(204, 9, new TFormula("=C204*H204"));

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
            xls.SetCellValue(205, 3, 19);
            xls.SetCellValue(205, 5, 1);
            xls.SetCellValue(205, 6, 1);
            xls.SetCellValue(205, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(205, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(205, 8, xls.AddFormat(fmt));
            xls.SetCellValue(205, 8, new TFormula("=(    1/ IF(E205<>1,VLOOKUP(E205,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) )  "
            + "    *      IF(F205<>1,VLOOKUP(F205,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)  "
            + "    * IF(G205<>1,VLOOKUP(G205,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(205, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(205, 9, xls.AddFormat(fmt));
            xls.SetCellValue(205, 9, new TFormula("=C205*H205"));

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
            xls.SetCellValue(206, 3, 13471);
            xls.SetCellValue(206, 5, new TFormula("='Gral Conf. Summary'!$H$33"));
            xls.SetCellValue(206, 6, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(206, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(206, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(206, 8, xls.AddFormat(fmt));
            xls.SetCellValue(206, 8, new TFormula("=(    1/ IF(E206<>1,VLOOKUP(E206,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) )  "
            + "    *      IF(F206<>1,VLOOKUP(F206,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)  "
            + "    * IF(G206<>1,VLOOKUP(G206,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(206, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(206, 9, xls.AddFormat(fmt));
            xls.SetCellValue(206, 9, new TFormula("=C206*H206"));

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
            xls.SetCellValue(207, 3, 8.8);
            xls.SetCellValue(207, 5, 1);
            xls.SetCellValue(207, 6, 1);
            xls.SetCellValue(207, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(207, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(207, 8, xls.AddFormat(fmt));
            xls.SetCellValue(207, 8, new TFormula("=(    1/ IF(E207<>1,VLOOKUP(E207,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) )  "
            + "    *      IF(F207<>1,VLOOKUP(F207,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)  "
            + "    * IF(G207<>1,VLOOKUP(G207,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(207, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(207, 9, xls.AddFormat(fmt));
            xls.SetCellValue(207, 9, new TFormula("=C207*H207"));

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
            xls.SetCellValue(208, 3, 0);
            xls.SetCellValue(208, 5, new TFormula("='Gral Conf. Summary'!$H$33"));
            xls.SetCellValue(208, 6, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(208, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(208, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(208, 8, xls.AddFormat(fmt));
            xls.SetCellValue(208, 8, new TFormula("=(    1/ IF(E208<>1,VLOOKUP(E208,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) )  "
            + "    *      IF(F208<>1,VLOOKUP(F208,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)  "
            + "    * IF(G208<>1,VLOOKUP(G208,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(208, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(208, 9, xls.AddFormat(fmt));
            xls.SetCellValue(208, 9, new TFormula("=C208*H208"));

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
            xls.SetCellValue(209, 3, 0);
            xls.SetCellValue(209, 5, 1);
            xls.SetCellValue(209, 6, 1);
            xls.SetCellValue(209, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(209, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(209, 8, xls.AddFormat(fmt));
            xls.SetCellValue(209, 8, new TFormula("=(    1/ IF(E209<>1,VLOOKUP(E209,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) )  "
            + "    *      IF(F209<>1,VLOOKUP(F209,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)  "
            + "    * IF(G209<>1,VLOOKUP(G209,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(209, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(209, 9, xls.AddFormat(fmt));
            xls.SetCellValue(209, 9, new TFormula("=C209*H209"));

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
            xls.SetCellValue(210, 3, 362);
            xls.SetCellValue(210, 5, new TFormula("='Gral Conf. Summary'!$H$33"));
            xls.SetCellValue(210, 6, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(210, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(210, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(210, 8, xls.AddFormat(fmt));
            xls.SetCellValue(210, 8, new TFormula("=(    1/ IF(E210<>1,VLOOKUP(E210,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) )  "
            + "    *      IF(F210<>1,VLOOKUP(F210,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)  "
            + "    * IF(G210<>1,VLOOKUP(G210,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(210, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(210, 9, xls.AddFormat(fmt));
            xls.SetCellValue(210, 9, new TFormula("=C210*H210"));

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
            xls.SetCellValue(211, 3, 1.3);
            xls.SetCellValue(211, 5, 1);
            xls.SetCellValue(211, 6, 1);
            xls.SetCellValue(211, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(211, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(211, 8, xls.AddFormat(fmt));
            xls.SetCellValue(211, 8, new TFormula("=(    1/ IF(E211<>1,VLOOKUP(E211,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) )  "
            + "    *      IF(F211<>1,VLOOKUP(F211,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)  "
            + "    * IF(G211<>1,VLOOKUP(G211,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(211, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(211, 9, xls.AddFormat(fmt));
            xls.SetCellValue(211, 9, new TFormula("=C211*H211"));

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
            xls.SetCellValue(212, 3, 1328);
            xls.SetCellValue(212, 5, new TFormula("='Gral Conf. Summary'!$H$33"));
            xls.SetCellValue(212, 6, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(212, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(212, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(212, 8, xls.AddFormat(fmt));
            xls.SetCellValue(212, 8, new TFormula("=(    1/ IF(E212<>1,VLOOKUP(E212,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) )  "
            + "    *      IF(F212<>1,VLOOKUP(F212,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)  "
            + "    * IF(G212<>1,VLOOKUP(G212,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(212, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(212, 9, xls.AddFormat(fmt));
            xls.SetCellValue(212, 9, new TFormula("=C212*H212"));

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
            xls.SetCellValue(213, 3, 2.7);
            xls.SetCellValue(213, 5, 1);
            xls.SetCellValue(213, 6, 1);
            xls.SetCellValue(213, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(213, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(213, 8, xls.AddFormat(fmt));
            xls.SetCellValue(213, 8, new TFormula("=(    1/ IF(E213<>1,VLOOKUP(E213,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) )  "
            + "    *      IF(F213<>1,VLOOKUP(F213,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)  "
            + "    * IF(G213<>1,VLOOKUP(G213,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(213, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(213, 9, xls.AddFormat(fmt));
            xls.SetCellValue(213, 9, new TFormula("=C213*H213"));

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
            xls.SetCellValue(214, 3, 72.16);
            xls.SetCellValue(214, 5, new TFormula("='Gral Conf. Summary'!$H$33"));
            xls.SetCellValue(214, 6, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(214, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(214, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(214, 8, xls.AddFormat(fmt));
            xls.SetCellValue(214, 8, new TFormula("=(    1/ IF(E214<>1,VLOOKUP(E214,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) )  "
            + "    *      IF(F214<>1,VLOOKUP(F214,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)  "
            + "    * IF(G214<>1,VLOOKUP(G214,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(214, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(214, 9, xls.AddFormat(fmt));
            xls.SetCellValue(214, 9, new TFormula("=C214*H214"));

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
            xls.SetCellValue(215, 3, 1.03);
            xls.SetCellValue(215, 5, 1);
            xls.SetCellValue(215, 6, 1);
            xls.SetCellValue(215, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(215, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(215, 8, xls.AddFormat(fmt));
            xls.SetCellValue(215, 8, new TFormula("=(    1/ IF(E215<>1,VLOOKUP(E215,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) )  "
            + "    *      IF(F215<>1,VLOOKUP(F215,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)  "
            + "    * IF(G215<>1,VLOOKUP(G215,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(215, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(215, 9, xls.AddFormat(fmt));
            xls.SetCellValue(215, 9, new TFormula("=C215*H215"));

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
            xls.SetCellValue(216, 3, 302);
            xls.SetCellValue(216, 5, new TFormula("='Gral Conf. Summary'!$H$33"));
            xls.SetCellValue(216, 6, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(216, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(216, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(216, 8, xls.AddFormat(fmt));
            xls.SetCellValue(216, 8, new TFormula("=(    1/ IF(E216<>1,VLOOKUP(E216,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) )  "
            + "    *      IF(F216<>1,VLOOKUP(F216,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)  "
            + "    * IF(G216<>1,VLOOKUP(G216,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(216, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(216, 9, xls.AddFormat(fmt));
            xls.SetCellValue(216, 9, new TFormula("=C216*H216"));

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
            xls.SetCellValue(217, 3, 1.35);
            xls.SetCellValue(217, 5, 1);
            xls.SetCellValue(217, 6, 1);
            xls.SetCellValue(217, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(217, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(217, 8, xls.AddFormat(fmt));
            xls.SetCellValue(217, 8, new TFormula("=(    1/ IF(E217<>1,VLOOKUP(E217,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) )  "
            + "    *      IF(F217<>1,VLOOKUP(F217,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)  "
            + "    * IF(G217<>1,VLOOKUP(G217,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(217, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(217, 9, xls.AddFormat(fmt));
            xls.SetCellValue(217, 9, new TFormula("=C217*H217"));

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
            xls.SetCellValue(218, 3, 228);
            xls.SetCellValue(218, 5, new TFormula("='Gral Conf. Summary'!$H$33"));
            xls.SetCellValue(218, 6, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(218, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(218, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(218, 8, xls.AddFormat(fmt));
            xls.SetCellValue(218, 8, new TFormula("=(    1/ IF(E218<>1,VLOOKUP(E218,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) )  "
            + "    *      IF(F218<>1,VLOOKUP(F218,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)  "
            + "    * IF(G218<>1,VLOOKUP(G218,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(218, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(218, 9, xls.AddFormat(fmt));
            xls.SetCellValue(218, 9, new TFormula("=C218*H218"));

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
            xls.SetCellValue(219, 3, 6.3);
            xls.SetCellValue(219, 5, 1);
            xls.SetCellValue(219, 6, 1);
            xls.SetCellValue(219, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(219, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(219, 8, xls.AddFormat(fmt));
            xls.SetCellValue(219, 8, new TFormula("=(    1/ IF(E219<>1,VLOOKUP(E219,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) )  "
            + "    *      IF(F219<>1,VLOOKUP(F219,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)  "
            + "    * IF(G219<>1,VLOOKUP(G219,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(219, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(219, 9, xls.AddFormat(fmt));
            xls.SetCellValue(219, 9, new TFormula("=C219*H219"));

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
            xls.SetCellValue(220, 3, 0);
            xls.SetCellValue(220, 5, new TFormula("='Gral Conf. Summary'!$H$33"));
            xls.SetCellValue(220, 6, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(220, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(220, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(220, 8, xls.AddFormat(fmt));
            xls.SetCellValue(220, 8, new TFormula("=(    1/ IF(E220<>1,VLOOKUP(E220,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) )  "
            + "    *      IF(F220<>1,VLOOKUP(F220,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)  "
            + "    * IF(G220<>1,VLOOKUP(G220,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(220, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(220, 9, xls.AddFormat(fmt));
            xls.SetCellValue(220, 9, new TFormula("=C220*H220"));

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
            xls.SetCellValue(221, 3, 1.9);
            xls.SetCellValue(221, 5, 1);
            xls.SetCellValue(221, 6, 1);
            xls.SetCellValue(221, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(221, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(221, 8, xls.AddFormat(fmt));
            xls.SetCellValue(221, 8, new TFormula("=(    1/ IF(E221<>1,VLOOKUP(E221,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) )  "
            + "    *      IF(F221<>1,VLOOKUP(F221,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)  "
            + "    * IF(G221<>1,VLOOKUP(G221,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(221, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(221, 9, xls.AddFormat(fmt));
            xls.SetCellValue(221, 9, new TFormula("=C221*H221"));

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

            fmt = xls.GetCellVisibleFormatDef(222, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(222, 9, xls.AddFormat(fmt));

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
            xls.SetCellValue(223, 3, 7946);
            xls.SetCellValue(223, 5, new TFormula("='Gral Conf. Summary'!$H$33"));
            xls.SetCellValue(223, 6, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(223, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(223, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(223, 8, xls.AddFormat(fmt));
            xls.SetCellValue(223, 8, new TFormula("=(    1/ IF(E223<>1,VLOOKUP(E223,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) )  "
            + "    *      IF(F223<>1,VLOOKUP(F223,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)  "
            + "    * IF(G223<>1,VLOOKUP(G223,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(223, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(223, 9, xls.AddFormat(fmt));
            xls.SetCellValue(223, 9, new TFormula("=C223*H223"));

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
            xls.SetCellValue(224, 3, 7.5);
            xls.SetCellValue(224, 5, 1);
            xls.SetCellValue(224, 6, 1);
            xls.SetCellValue(224, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(224, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(224, 8, xls.AddFormat(fmt));
            xls.SetCellValue(224, 8, new TFormula("=(    1/ IF(E224<>1,VLOOKUP(E224,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) )  "
            + "    *      IF(F224<>1,VLOOKUP(F224,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)  "
            + "    * IF(G224<>1,VLOOKUP(G224,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(224, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(224, 9, xls.AddFormat(fmt));
            xls.SetCellValue(224, 9, new TFormula("=C224*H224"));

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
            xls.SetCellValue(225, 3, 0);
            xls.SetCellValue(225, 5, new TFormula("='Gral Conf. Summary'!$H$33"));
            xls.SetCellValue(225, 6, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(225, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(225, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(225, 8, xls.AddFormat(fmt));
            xls.SetCellValue(225, 8, new TFormula("=(    1/ IF(E225<>1,VLOOKUP(E225,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) )  "
            + "    *      IF(F225<>1,VLOOKUP(F225,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)  "
            + "    * IF(G225<>1,VLOOKUP(G225,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(225, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(225, 9, xls.AddFormat(fmt));
            xls.SetCellValue(225, 9, new TFormula("=C225*H225"));

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
            xls.SetCellValue(226, 3, 0.1);
            xls.SetCellValue(226, 5, 1);
            xls.SetCellValue(226, 6, 1);
            xls.SetCellValue(226, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(226, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(226, 8, xls.AddFormat(fmt));
            xls.SetCellValue(226, 8, new TFormula("=(    1/ IF(E226<>1,VLOOKUP(E226,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) )  "
            + "    *      IF(F226<>1,VLOOKUP(F226,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)  "
            + "    * IF(G226<>1,VLOOKUP(G226,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(226, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(226, 9, xls.AddFormat(fmt));
            xls.SetCellValue(226, 9, new TFormula("=C226*H226"));

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
            xls.SetCellValue(227, 3, 6565);
            xls.SetCellValue(227, 5, new TFormula("='Gral Conf. Summary'!$H$33"));
            xls.SetCellValue(227, 6, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(227, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(227, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(227, 8, xls.AddFormat(fmt));
            xls.SetCellValue(227, 8, new TFormula("=(    1/ IF(E227<>1,VLOOKUP(E227,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) )  "
            + "    *      IF(F227<>1,VLOOKUP(F227,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)  "
            + "    * IF(G227<>1,VLOOKUP(G227,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(227, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(227, 9, xls.AddFormat(fmt));
            xls.SetCellValue(227, 9, new TFormula("=C227*H227"));

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
            xls.SetCellValue(228, 3, 8.78);
            xls.SetCellValue(228, 5, 1);
            xls.SetCellValue(228, 6, 1);
            xls.SetCellValue(228, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(228, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(228, 8, xls.AddFormat(fmt));
            xls.SetCellValue(228, 8, new TFormula("=(    1/ IF(E228<>1,VLOOKUP(E228,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) )  "
            + "    *      IF(F228<>1,VLOOKUP(F228,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)  "
            + "    * IF(G228<>1,VLOOKUP(G228,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(228, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(228, 9, xls.AddFormat(fmt));
            xls.SetCellValue(228, 9, new TFormula("=C228*H228"));

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
            xls.SetCellValue(229, 3, 10236);
            xls.SetCellValue(229, 5, new TFormula("='Gral Conf. Summary'!$H$33"));
            xls.SetCellValue(229, 6, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(229, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(229, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(229, 8, xls.AddFormat(fmt));
            xls.SetCellValue(229, 8, new TFormula("=(    1/ IF(E229<>1,VLOOKUP(E229,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) )  "
            + "    *      IF(F229<>1,VLOOKUP(F229,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)  "
            + "    * IF(G229<>1,VLOOKUP(G229,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(229, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(229, 9, xls.AddFormat(fmt));
            xls.SetCellValue(229, 9, new TFormula("=C229*H229"));

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
            xls.SetCellValue(230, 3, 8.77);
            xls.SetCellValue(230, 5, 1);
            xls.SetCellValue(230, 6, 1);
            xls.SetCellValue(230, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(230, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(230, 8, xls.AddFormat(fmt));
            xls.SetCellValue(230, 8, new TFormula("=(    1/ IF(E230<>1,VLOOKUP(E230,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) )  "
            + "    *      IF(F230<>1,VLOOKUP(F230,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)  "
            + "    * IF(G230<>1,VLOOKUP(G230,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(230, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(230, 9, xls.AddFormat(fmt));
            xls.SetCellValue(230, 9, new TFormula("=C230*H230"));

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
            xls.SetCellValue(231, 3, 2389);
            xls.SetCellValue(231, 5, new TFormula("='Gral Conf. Summary'!$H$33"));
            xls.SetCellValue(231, 6, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(231, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(231, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(231, 8, xls.AddFormat(fmt));
            xls.SetCellValue(231, 8, new TFormula("=(    1/ IF(E231<>1,VLOOKUP(E231,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) )  "
            + "    *      IF(F231<>1,VLOOKUP(F231,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)  "
            + "    * IF(G231<>1,VLOOKUP(G231,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(231, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(231, 9, xls.AddFormat(fmt));
            xls.SetCellValue(231, 9, new TFormula("=C231*H231"));

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
            xls.SetCellValue(232, 3, 7.47);
            xls.SetCellValue(232, 5, 1);
            xls.SetCellValue(232, 6, 1);
            xls.SetCellValue(232, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(232, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(232, 8, xls.AddFormat(fmt));
            xls.SetCellValue(232, 8, new TFormula("=(    1/ IF(E232<>1,VLOOKUP(E232,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) )  "
            + "    *      IF(F232<>1,VLOOKUP(F232,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)  "
            + "    * IF(G232<>1,VLOOKUP(G232,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(232, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(232, 9, xls.AddFormat(fmt));
            xls.SetCellValue(232, 9, new TFormula("=C232*H232"));

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
            xls.SetCellValue(233, 3, 392);
            xls.SetCellValue(233, 5, new TFormula("='Gral Conf. Summary'!$H$33"));
            xls.SetCellValue(233, 6, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(233, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(233, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(233, 8, xls.AddFormat(fmt));
            xls.SetCellValue(233, 8, new TFormula("=(    1/ IF(E233<>1,VLOOKUP(E233,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) )  "
            + "    *      IF(F233<>1,VLOOKUP(F233,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)  "
            + "    * IF(G233<>1,VLOOKUP(G233,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(233, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(233, 9, xls.AddFormat(fmt));
            xls.SetCellValue(233, 9, new TFormula("=C233*H233"));

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
            xls.SetCellValue(234, 3, 5.16);
            xls.SetCellValue(234, 5, 1);
            xls.SetCellValue(234, 6, 1);
            xls.SetCellValue(234, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(234, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(234, 8, xls.AddFormat(fmt));
            xls.SetCellValue(234, 8, new TFormula("=(    1/ IF(E234<>1,VLOOKUP(E234,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) )  "
            + "    *      IF(F234<>1,VLOOKUP(F234,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)  "
            + "    * IF(G234<>1,VLOOKUP(G234,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(234, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(234, 9, xls.AddFormat(fmt));
            xls.SetCellValue(234, 9, new TFormula("=C234*H234"));

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
            xls.SetCellValue(235, 3, 530);
            xls.SetCellValue(235, 5, new TFormula("='Gral Conf. Summary'!$H$33"));
            xls.SetCellValue(235, 6, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(235, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(235, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(235, 8, xls.AddFormat(fmt));
            xls.SetCellValue(235, 8, new TFormula("=(    1/ IF(E235<>1,VLOOKUP(E235,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) )  "
            + "    *      IF(F235<>1,VLOOKUP(F235,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)  "
            + "    * IF(G235<>1,VLOOKUP(G235,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(235, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(235, 9, xls.AddFormat(fmt));
            xls.SetCellValue(235, 9, new TFormula("=C235*H235"));

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
            xls.SetCellValue(236, 3, 6.13);
            xls.SetCellValue(236, 5, 1);
            xls.SetCellValue(236, 6, 1);
            xls.SetCellValue(236, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(236, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(236, 8, xls.AddFormat(fmt));
            xls.SetCellValue(236, 8, new TFormula("=(    1/ IF(E236<>1,VLOOKUP(E236,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) )  "
            + "    *      IF(F236<>1,VLOOKUP(F236,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)  "
            + "    * IF(G236<>1,VLOOKUP(G236,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(236, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(236, 9, xls.AddFormat(fmt));
            xls.SetCellValue(236, 9, new TFormula("=C236*H236"));

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
            xls.SetCellValue(237, 3, 227);
            xls.SetCellValue(237, 5, new TFormula("='Gral Conf. Summary'!$H$33"));
            xls.SetCellValue(237, 6, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(237, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(237, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(237, 8, xls.AddFormat(fmt));
            xls.SetCellValue(237, 8, new TFormula("=(    1/ IF(E237<>1,VLOOKUP(E237,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) )  "
            + "    *      IF(F237<>1,VLOOKUP(F237,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)  "
            + "    * IF(G237<>1,VLOOKUP(G237,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(237, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(237, 9, xls.AddFormat(fmt));
            xls.SetCellValue(237, 9, new TFormula("=C237*H237"));

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
            xls.SetCellValue(238, 3, 5.3);
            xls.SetCellValue(238, 5, 1);
            xls.SetCellValue(238, 6, 1);
            xls.SetCellValue(238, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(238, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(238, 8, xls.AddFormat(fmt));
            xls.SetCellValue(238, 8, new TFormula("=(    1/ IF(E238<>1,VLOOKUP(E238,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) )  "
            + "    *      IF(F238<>1,VLOOKUP(F238,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)  "
            + "    * IF(G238<>1,VLOOKUP(G238,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(238, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(238, 9, xls.AddFormat(fmt));
            xls.SetCellValue(238, 9, new TFormula("=C238*H238"));

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
            xls.SetCellValue(239, 3, 0);
            xls.SetCellValue(239, 5, new TFormula("='Gral Conf. Summary'!$H$33"));
            xls.SetCellValue(239, 6, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(239, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(239, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(239, 8, xls.AddFormat(fmt));
            xls.SetCellValue(239, 8, new TFormula("=(    1/ IF(E239<>1,VLOOKUP(E239,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) )  "
            + "    *      IF(F239<>1,VLOOKUP(F239,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)  "
            + "    * IF(G239<>1,VLOOKUP(G239,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(239, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(239, 9, xls.AddFormat(fmt));
            xls.SetCellValue(239, 9, new TFormula("=C239*H239"));

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
            xls.SetCellValue(240, 3, 0);
            xls.SetCellValue(240, 5, 1);
            xls.SetCellValue(240, 6, 1);
            xls.SetCellValue(240, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(240, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(240, 8, xls.AddFormat(fmt));
            xls.SetCellValue(240, 8, new TFormula("=(    1/ IF(E240<>1,VLOOKUP(E240,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) )  "
            + "    *      IF(F240<>1,VLOOKUP(F240,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)  "
            + "    * IF(G240<>1,VLOOKUP(G240,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(240, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(240, 9, xls.AddFormat(fmt));
            xls.SetCellValue(240, 9, new TFormula("=C240*H240"));

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
            xls.SetCellValue(241, 3, 442);
            xls.SetCellValue(241, 5, new TFormula("='Gral Conf. Summary'!$H$33"));
            xls.SetCellValue(241, 6, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(241, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(241, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(241, 8, xls.AddFormat(fmt));
            xls.SetCellValue(241, 8, new TFormula("=(    1/ IF(E241<>1,VLOOKUP(E241,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) )  "
            + "    *      IF(F241<>1,VLOOKUP(F241,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)  "
            + "    * IF(G241<>1,VLOOKUP(G241,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(241, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(241, 9, xls.AddFormat(fmt));
            xls.SetCellValue(241, 9, new TFormula("=C241*H241"));

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
            xls.SetCellValue(242, 3, 9.5);
            xls.SetCellValue(242, 5, 1);
            xls.SetCellValue(242, 6, 1);
            xls.SetCellValue(242, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(242, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(242, 8, xls.AddFormat(fmt));
            xls.SetCellValue(242, 8, new TFormula("=(    1/ IF(E242<>1,VLOOKUP(E242,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) )  "
            + "    *      IF(F242<>1,VLOOKUP(F242,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)  "
            + "    * IF(G242<>1,VLOOKUP(G242,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(242, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(242, 9, xls.AddFormat(fmt));
            xls.SetCellValue(242, 9, new TFormula("=C242*H242"));

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
            xls.SetCellValue(243, 3, 0);
            xls.SetCellValue(243, 5, new TFormula("='Gral Conf. Summary'!$H$33"));
            xls.SetCellValue(243, 6, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(243, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(243, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(243, 8, xls.AddFormat(fmt));
            xls.SetCellValue(243, 8, new TFormula("=(    1/ IF(E243<>1,VLOOKUP(E243,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) )  "
            + "    *      IF(F243<>1,VLOOKUP(F243,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)  "
            + "    * IF(G243<>1,VLOOKUP(G243,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(243, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(243, 9, xls.AddFormat(fmt));
            xls.SetCellValue(243, 9, new TFormula("=C243*H243"));

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
            xls.SetCellValue(244, 3, 0.1);
            xls.SetCellValue(244, 5, 1);
            xls.SetCellValue(244, 6, 1);
            xls.SetCellValue(244, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(244, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(244, 8, xls.AddFormat(fmt));
            xls.SetCellValue(244, 8, new TFormula("=(    1/ IF(E244<>1,VLOOKUP(E244,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) )  "
            + "    *      IF(F244<>1,VLOOKUP(F244,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)  "
            + "    * IF(G244<>1,VLOOKUP(G244,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(244, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(244, 9, xls.AddFormat(fmt));
            xls.SetCellValue(244, 9, new TFormula("=C244*H244"));

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
            xls.SetCellValue(245, 3, 25522);
            xls.SetCellValue(245, 5, new TFormula("='Gral Conf. Summary'!$H$33"));
            xls.SetCellValue(245, 6, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(245, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(245, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(245, 8, xls.AddFormat(fmt));
            xls.SetCellValue(245, 8, new TFormula("=(    1/ IF(E245<>1,VLOOKUP(E245,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) )  "
            + "    *      IF(F245<>1,VLOOKUP(F245,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)  "
            + "    * IF(G245<>1,VLOOKUP(G245,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(245, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(245, 9, xls.AddFormat(fmt));
            xls.SetCellValue(245, 9, new TFormula("=C245*H245"));

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
            xls.SetCellValue(246, 3, 8.3);
            xls.SetCellValue(246, 5, 1);
            xls.SetCellValue(246, 6, 1);
            xls.SetCellValue(246, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(246, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(246, 8, xls.AddFormat(fmt));
            xls.SetCellValue(246, 8, new TFormula("=(    1/ IF(E246<>1,VLOOKUP(E246,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) )  "
            + "    *      IF(F246<>1,VLOOKUP(F246,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)  "
            + "    * IF(G246<>1,VLOOKUP(G246,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(246, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(246, 9, xls.AddFormat(fmt));
            xls.SetCellValue(246, 9, new TFormula("=C246*H246"));

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
            xls.SetCellValue(247, 3, 1521);
            xls.SetCellValue(247, 5, new TFormula("='Gral Conf. Summary'!$H$33"));
            xls.SetCellValue(247, 6, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(247, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(247, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(247, 8, xls.AddFormat(fmt));
            xls.SetCellValue(247, 8, new TFormula("=(    1/ IF(E247<>1,VLOOKUP(E247,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) )  "
            + "    *      IF(F247<>1,VLOOKUP(F247,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)  "
            + "    * IF(G247<>1,VLOOKUP(G247,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(247, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(247, 9, xls.AddFormat(fmt));
            xls.SetCellValue(247, 9, new TFormula("=C247*H247"));

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
            xls.SetCellValue(248, 3, 3.43);
            xls.SetCellValue(248, 5, 1);
            xls.SetCellValue(248, 6, 1);
            xls.SetCellValue(248, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(248, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(248, 8, xls.AddFormat(fmt));
            xls.SetCellValue(248, 8, new TFormula("=(    1/ IF(E248<>1,VLOOKUP(E248,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) )  "
            + "    *      IF(F248<>1,VLOOKUP(F248,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)  "
            + "    * IF(G248<>1,VLOOKUP(G248,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(248, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(248, 9, xls.AddFormat(fmt));
            xls.SetCellValue(248, 9, new TFormula("=C248*H248"));

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
            xls.SetCellValue(249, 3, 228);
            xls.SetCellValue(249, 5, new TFormula("='Gral Conf. Summary'!$H$33"));
            xls.SetCellValue(249, 6, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(249, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(249, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(249, 8, xls.AddFormat(fmt));
            xls.SetCellValue(249, 8, new TFormula("=(    1/ IF(E249<>1,VLOOKUP(E249,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) )  "
            + "    *      IF(F249<>1,VLOOKUP(F249,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)  "
            + "    * IF(G249<>1,VLOOKUP(G249,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(249, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(249, 9, xls.AddFormat(fmt));
            xls.SetCellValue(249, 9, new TFormula("=C249*H249"));

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
            xls.SetCellValue(250, 3, 2.91);
            xls.SetCellValue(250, 5, 1);
            xls.SetCellValue(250, 6, 1);
            xls.SetCellValue(250, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(250, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(250, 8, xls.AddFormat(fmt));
            xls.SetCellValue(250, 8, new TFormula("=(    1/ IF(E250<>1,VLOOKUP(E250,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) )  "
            + "    *      IF(F250<>1,VLOOKUP(F250,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)  "
            + "    * IF(G250<>1,VLOOKUP(G250,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(250, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(250, 9, xls.AddFormat(fmt));
            xls.SetCellValue(250, 9, new TFormula("=C250*H250"));

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
            xls.SetCellValue(251, 3, 50);
            xls.SetCellValue(251, 5, new TFormula("='Gral Conf. Summary'!$H$33"));
            xls.SetCellValue(251, 6, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(251, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(251, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(251, 8, xls.AddFormat(fmt));
            xls.SetCellValue(251, 8, new TFormula("=(    1/ IF(E251<>1,VLOOKUP(E251,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) )  "
            + "    *      IF(F251<>1,VLOOKUP(F251,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)  "
            + "    * IF(G251<>1,VLOOKUP(G251,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(251, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(251, 9, xls.AddFormat(fmt));
            xls.SetCellValue(251, 9, new TFormula("=C251*H251"));

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
            xls.SetCellValue(252, 3, 1.4);
            xls.SetCellValue(252, 5, 1);
            xls.SetCellValue(252, 6, 1);
            xls.SetCellValue(252, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(252, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(252, 8, xls.AddFormat(fmt));
            xls.SetCellValue(252, 8, new TFormula("=(    1/ IF(E252<>1,VLOOKUP(E252,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) )  "
            + "    *      IF(F252<>1,VLOOKUP(F252,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)  "
            + "    * IF(G252<>1,VLOOKUP(G252,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(252, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(252, 9, xls.AddFormat(fmt));
            xls.SetCellValue(252, 9, new TFormula("=C252*H252"));

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
            xls.SetCellValue(253, 3, 0);
            xls.SetCellValue(253, 5, new TFormula("='Gral Conf. Summary'!$H$33"));
            xls.SetCellValue(253, 6, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(253, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(253, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(253, 8, xls.AddFormat(fmt));
            xls.SetCellValue(253, 8, new TFormula("=(    1/ IF(E253<>1,VLOOKUP(E253,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) )  "
            + "    *      IF(F253<>1,VLOOKUP(F253,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)  "
            + "    * IF(G253<>1,VLOOKUP(G253,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(253, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(253, 9, xls.AddFormat(fmt));
            xls.SetCellValue(253, 9, new TFormula("=C253*H253"));

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
            xls.SetCellValue(254, 3, 0.1);
            xls.SetCellValue(254, 5, 1);
            xls.SetCellValue(254, 6, 1);
            xls.SetCellValue(254, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(254, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(254, 8, xls.AddFormat(fmt));
            xls.SetCellValue(254, 8, new TFormula("=(    1/ IF(E254<>1,VLOOKUP(E254,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) )  "
            + "    *      IF(F254<>1,VLOOKUP(F254,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)  "
            + "    * IF(G254<>1,VLOOKUP(G254,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(254, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(254, 9, xls.AddFormat(fmt));
            xls.SetCellValue(254, 9, new TFormula("=C254*H254"));

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
            xls.SetCellValue(255, 3, 75);
            xls.SetCellValue(255, 5, new TFormula("='Gral Conf. Summary'!$H$33"));
            xls.SetCellValue(255, 6, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(255, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(255, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(255, 8, xls.AddFormat(fmt));
            xls.SetCellValue(255, 8, new TFormula("=(    1/ IF(E255<>1,VLOOKUP(E255,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) )  "
            + "    *      IF(F255<>1,VLOOKUP(F255,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)  "
            + "    * IF(G255<>1,VLOOKUP(G255,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(255, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(255, 9, xls.AddFormat(fmt));
            xls.SetCellValue(255, 9, new TFormula("=C255*H255"));

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
            xls.SetCellValue(256, 3, 1.5);
            xls.SetCellValue(256, 5, 1);
            xls.SetCellValue(256, 6, 1);
            xls.SetCellValue(256, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(256, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(256, 8, xls.AddFormat(fmt));
            xls.SetCellValue(256, 8, new TFormula("=(    1/ IF(E256<>1,VLOOKUP(E256,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) )  "
            + "    *      IF(F256<>1,VLOOKUP(F256,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)  "
            + "    * IF(G256<>1,VLOOKUP(G256,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(256, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(256, 9, xls.AddFormat(fmt));
            xls.SetCellValue(256, 9, new TFormula("=C256*H256"));

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

            fmt = xls.GetCellVisibleFormatDef(257, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(257, 9, xls.AddFormat(fmt));

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

            fmt = xls.GetCellVisibleFormatDef(258, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(258, 9, xls.AddFormat(fmt));

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
            xls.SetCellValue(259, 3, 2236);
            xls.SetCellValue(259, 5, new TFormula("='Gral Conf. Summary'!$H$33"));
            xls.SetCellValue(259, 6, 1);
            xls.SetCellValue(259, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(259, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(259, 8, xls.AddFormat(fmt));
            xls.SetCellValue(259, 8, new TFormula("=( 1/  IF(E259<>1,VLOOKUP(E259,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)  )   "
            + "  *       IF(F259<>1,VLOOKUP(F259,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) * IF(G259<>1,VLOOKUP(G259,'Gral"
            + " Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(259, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(259, 9, xls.AddFormat(fmt));
            xls.SetCellValue(259, 9, new TFormula("=C259*H259"));

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
            xls.SetCellValue(260, 3, 0);
            xls.SetCellValue(260, 5, new TFormula("='Gral Conf. Summary'!$H$33"));
            xls.SetCellValue(260, 6, 1);
            xls.SetCellValue(260, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(260, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(260, 8, xls.AddFormat(fmt));
            xls.SetCellValue(260, 8, new TFormula("=( 1/  IF(E260<>1,VLOOKUP(E260,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)  )   "
            + "  *       IF(F260<>1,VLOOKUP(F260,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) * IF(G260<>1,VLOOKUP(G260,'Gral"
            + " Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(260, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(260, 9, xls.AddFormat(fmt));
            xls.SetCellValue(260, 9, new TFormula("=C260*H260"));

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
            xls.SetCellValue(261, 3, 0);
            xls.SetCellValue(261, 5, new TFormula("='Gral Conf. Summary'!$H$33"));
            xls.SetCellValue(261, 6, 1);
            xls.SetCellValue(261, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(261, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(261, 8, xls.AddFormat(fmt));
            xls.SetCellValue(261, 8, new TFormula("=( 1/  IF(E261<>1,VLOOKUP(E261,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)  )   "
            + "  *       IF(F261<>1,VLOOKUP(F261,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) * IF(G261<>1,VLOOKUP(G261,'Gral"
            + " Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(261, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(261, 9, xls.AddFormat(fmt));
            xls.SetCellValue(261, 9, new TFormula("=C261*H261"));

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
            xls.SetCellValue(262, 3, 0);
            xls.SetCellValue(262, 5, new TFormula("='Gral Conf. Summary'!$H$33"));
            xls.SetCellValue(262, 6, 1);
            xls.SetCellValue(262, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(262, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(262, 8, xls.AddFormat(fmt));
            xls.SetCellValue(262, 8, new TFormula("=( 1/  IF(E262<>1,VLOOKUP(E262,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)  )   "
            + "  *       IF(F262<>1,VLOOKUP(F262,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) * IF(G262<>1,VLOOKUP(G262,'Gral"
            + " Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(262, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(262, 9, xls.AddFormat(fmt));
            xls.SetCellValue(262, 9, new TFormula("=C262*H262"));

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
            xls.SetCellValue(263, 3, 0);
            xls.SetCellValue(263, 5, new TFormula("='Gral Conf. Summary'!$H$33"));
            xls.SetCellValue(263, 6, 1);
            xls.SetCellValue(263, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(263, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(263, 8, xls.AddFormat(fmt));
            xls.SetCellValue(263, 8, new TFormula("=( 1/  IF(E263<>1,VLOOKUP(E263,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)  )   "
            + "  *       IF(F263<>1,VLOOKUP(F263,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) * IF(G263<>1,VLOOKUP(G263,'Gral"
            + " Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(263, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(263, 9, xls.AddFormat(fmt));
            xls.SetCellValue(263, 9, new TFormula("=C263*H263"));

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

            fmt = xls.GetCellVisibleFormatDef(264, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(264, 9, xls.AddFormat(fmt));

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
            xls.SetCellValue(265, 3, 36217);
            xls.SetCellValue(265, 5, new TFormula("='Gral Conf. Summary'!$H$33"));
            xls.SetCellValue(265, 6, 1);
            xls.SetCellValue(265, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(265, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(265, 8, xls.AddFormat(fmt));
            xls.SetCellValue(265, 8, new TFormula("=( 1/  IF(E265<>1,VLOOKUP(E265,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)  )   "
            + "  *       IF(F265<>1,VLOOKUP(F265,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) * IF(G265<>1,VLOOKUP(G265,'Gral"
            + " Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(265, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(265, 9, xls.AddFormat(fmt));
            xls.SetCellValue(265, 9, new TFormula("=C265*H265"));

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
            xls.SetCellValue(266, 3, 147.28);
            xls.SetCellValue(266, 5, new TFormula("='Gral Conf. Summary'!$H$33"));
            xls.SetCellValue(266, 6, 1);
            xls.SetCellValue(266, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(266, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(266, 8, xls.AddFormat(fmt));
            xls.SetCellValue(266, 8, new TFormula("=( 1/  IF(E266<>1,VLOOKUP(E266,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)  )   "
            + "  *       IF(F266<>1,VLOOKUP(F266,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) * IF(G266<>1,VLOOKUP(G266,'Gral"
            + " Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(266, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(266, 9, xls.AddFormat(fmt));
            xls.SetCellValue(266, 9, new TFormula("=C266*H266"));

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

            fmt = xls.GetCellVisibleFormatDef(267, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(267, 9, xls.AddFormat(fmt));
            xls.SetCellValue(267, 9, new TFormula("=C267*H267"));

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
            xls.SetCellValue(268, 3, 14);
            xls.SetCellValue(268, 5, 1);
            xls.SetCellValue(268, 6, 1);
            xls.SetCellValue(268, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(268, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(268, 8, xls.AddFormat(fmt));
            xls.SetCellValue(268, 8, new TFormula("= IF(E268<>1,VLOOKUP(E268,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)       *   "
            + "    IF(F268<>1,VLOOKUP(F268,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) * IF(G268<>1,VLOOKUP(G268,'Gral"
            + " Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(268, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(268, 9, xls.AddFormat(fmt));
            xls.SetCellValue(268, 9, new TFormula("=C268*H268"));

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
            xls.SetCellValue(269, 3, 5.37);
            xls.SetCellValue(269, 5, 1);
            xls.SetCellValue(269, 6, 1);
            xls.SetCellValue(269, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(269, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(269, 8, xls.AddFormat(fmt));
            xls.SetCellValue(269, 8, new TFormula("= IF(E269<>1,VLOOKUP(E269,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)       *   "
            + "    IF(F269<>1,VLOOKUP(F269,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) * IF(G269<>1,VLOOKUP(G269,'Gral"
            + " Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(269, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(269, 9, xls.AddFormat(fmt));
            xls.SetCellValue(269, 9, new TFormula("=C269*H269"));

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
            xls.SetCellValue(270, 3, 1.4);
            xls.SetCellValue(270, 5, 1);
            xls.SetCellValue(270, 6, 1);
            xls.SetCellValue(270, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(270, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(270, 8, xls.AddFormat(fmt));
            xls.SetCellValue(270, 8, new TFormula("= IF(E270<>1,VLOOKUP(E270,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)       *   "
            + "    IF(F270<>1,VLOOKUP(F270,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) * IF(G270<>1,VLOOKUP(G270,'Gral"
            + " Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(270, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(270, 9, xls.AddFormat(fmt));
            xls.SetCellValue(270, 9, new TFormula("=C270*H270"));

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
            xls.SetCellValue(271, 3, 899);
            xls.SetCellValue(271, 5, new TFormula("='Gral Conf. Summary'!$H$33"));
            xls.SetCellValue(271, 6, 1);
            xls.SetCellValue(271, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(271, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(271, 8, xls.AddFormat(fmt));
            xls.SetCellValue(271, 8, new TFormula("=( 1/  IF(E271<>1,VLOOKUP(E271,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)  )   "
            + "  *       IF(F271<>1,VLOOKUP(F271,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) * IF(G271<>1,VLOOKUP(G271,'Gral"
            + " Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(271, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(271, 9, xls.AddFormat(fmt));
            xls.SetCellValue(271, 9, new TFormula("=C271*H271"));

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

            fmt = xls.GetCellVisibleFormatDef(272, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(272, 9, xls.AddFormat(fmt));
            xls.SetCellValue(272, 9, new TFormula("=C272*H272"));

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

            fmt = xls.GetCellVisibleFormatDef(273, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(273, 9, xls.AddFormat(fmt));
            xls.SetCellValue(273, 9, new TFormula("=C273*H273"));

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

            fmt = xls.GetCellVisibleFormatDef(274, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(274, 9, xls.AddFormat(fmt));
            xls.SetCellValue(274, 9, new TFormula("=C274*H274"));

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
            xls.SetCellValue(275, 3, 79.78);
            xls.SetCellValue(275, 5, new TFormula("='Gral Conf. Summary'!$H$33"));
            xls.SetCellValue(275, 6, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(275, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(275, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(275, 8, xls.AddFormat(fmt));
            xls.SetCellValue(275, 8, new TFormula("=(    1/ IF(E275<>1,VLOOKUP(E275,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) )  "
            + "    *      IF(F275<>1,VLOOKUP(F275,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)  "
            + "    * IF(G275<>1,VLOOKUP(G275,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(275, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(275, 9, xls.AddFormat(fmt));
            xls.SetCellValue(275, 9, new TFormula("=C275*H275"));

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
            xls.SetCellValue(276, 3, 113.84);
            xls.SetCellValue(276, 5, new TFormula("='Gral Conf. Summary'!$H$33"));
            xls.SetCellValue(276, 6, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(276, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(276, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(276, 8, xls.AddFormat(fmt));
            xls.SetCellValue(276, 8, new TFormula("=(    1/ IF(E276<>1,VLOOKUP(E276,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) )  "
            + "    *      IF(F276<>1,VLOOKUP(F276,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)  "
            + "    * IF(G276<>1,VLOOKUP(G276,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(276, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(276, 9, xls.AddFormat(fmt));
            xls.SetCellValue(276, 9, new TFormula("=C276*H276"));

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
            xls.SetCellValue(277, 3, 172.32);
            xls.SetCellValue(277, 5, new TFormula("='Gral Conf. Summary'!$H$33"));
            xls.SetCellValue(277, 6, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(277, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(277, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(277, 8, xls.AddFormat(fmt));
            xls.SetCellValue(277, 8, new TFormula("=(    1/ IF(E277<>1,VLOOKUP(E277,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) )  "
            + "    *      IF(F277<>1,VLOOKUP(F277,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)  "
            + "    * IF(G277<>1,VLOOKUP(G277,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(277, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(277, 9, xls.AddFormat(fmt));
            xls.SetCellValue(277, 9, new TFormula("=C277*H277"));

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
            xls.SetCellValue(278, 3, 0);
            xls.SetCellValue(278, 5, new TFormula("='Gral Conf. Summary'!$H$33"));
            xls.SetCellValue(278, 6, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(278, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(278, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(278, 8, xls.AddFormat(fmt));
            xls.SetCellValue(278, 8, new TFormula("=(    1/ IF(E278<>1,VLOOKUP(E278,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) )  "
            + "    *      IF(F278<>1,VLOOKUP(F278,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)  "
            + "    * IF(G278<>1,VLOOKUP(G278,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(278, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(278, 9, xls.AddFormat(fmt));
            xls.SetCellValue(278, 9, new TFormula("=C278*H278"));

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

            fmt = xls.GetCellVisibleFormatDef(279, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(279, 9, xls.AddFormat(fmt));

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
            xls.SetCellValue(280, 3, 475.84);
            xls.SetCellValue(280, 5, new TFormula("='Gral Conf. Summary'!$H$33"));
            xls.SetCellValue(280, 6, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(280, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(280, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(280, 8, xls.AddFormat(fmt));
            xls.SetCellValue(280, 8, new TFormula("=(    1/ IF(E280<>1,VLOOKUP(E280,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) )  "
            + "    *      IF(F280<>1,VLOOKUP(F280,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)  "
            + "    * IF(G280<>1,VLOOKUP(G280,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(280, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(280, 9, xls.AddFormat(fmt));
            xls.SetCellValue(280, 9, new TFormula("=C280*H280"));

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
            xls.SetCellValue(281, 3, new TFormula("=C291/4"));
            xls.SetCellValue(281, 5, new TFormula("='Gral Conf. Summary'!$H$33"));
            xls.SetCellValue(281, 6, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(281, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(281, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(281, 8, xls.AddFormat(fmt));
            xls.SetCellValue(281, 8, new TFormula("=(    1/ IF(E281<>1,VLOOKUP(E281,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) )  "
            + "    *      IF(F281<>1,VLOOKUP(F281,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)  "
            + "    * IF(G281<>1,VLOOKUP(G281,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(281, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(281, 9, xls.AddFormat(fmt));
            xls.SetCellValue(281, 9, new TFormula("=C281*H281"));

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
            xls.SetCellValue(282, 3, 0);
            xls.SetCellValue(282, 5, new TFormula("='Gral Conf. Summary'!$H$33"));
            xls.SetCellValue(282, 6, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(282, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(282, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(282, 8, xls.AddFormat(fmt));
            xls.SetCellValue(282, 8, new TFormula("=(    1/ IF(E282<>1,VLOOKUP(E282,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) )  "
            + "    *      IF(F282<>1,VLOOKUP(F282,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)  "
            + "    * IF(G282<>1,VLOOKUP(G282,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(282, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(282, 9, xls.AddFormat(fmt));
            xls.SetCellValue(282, 9, new TFormula("=C282*H282"));

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

            fmt = xls.GetCellVisibleFormatDef(283, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(283, 9, xls.AddFormat(fmt));

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
            xls.SetCellValue(284, 3, 266.4);
            xls.SetCellValue(284, 5, new TFormula("='Gral Conf. Summary'!$H$33"));
            xls.SetCellValue(284, 6, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(284, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(284, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(284, 8, xls.AddFormat(fmt));
            xls.SetCellValue(284, 8, new TFormula("=(    1/ IF(E284<>1,VLOOKUP(E284,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) )  "
            + "    *      IF(F284<>1,VLOOKUP(F284,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)  "
            + "    * IF(G284<>1,VLOOKUP(G284,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(284, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(284, 9, xls.AddFormat(fmt));
            xls.SetCellValue(284, 9, new TFormula("=C284*H284"));

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
            xls.SetCellValue(285, 3, 142.12);
            xls.SetCellValue(285, 5, new TFormula("='Gral Conf. Summary'!$H$33"));
            xls.SetCellValue(285, 6, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(285, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(285, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(285, 8, xls.AddFormat(fmt));
            xls.SetCellValue(285, 8, new TFormula("=(    1/ IF(E285<>1,VLOOKUP(E285,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) )  "
            + "    *      IF(F285<>1,VLOOKUP(F285,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)  "
            + "    * IF(G285<>1,VLOOKUP(G285,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(285, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(285, 9, xls.AddFormat(fmt));
            xls.SetCellValue(285, 9, new TFormula("=C285*H285"));

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
            xls.SetCellValue(286, 3, 1817.6);
            xls.SetCellValue(286, 5, new TFormula("='Gral Conf. Summary'!$H$33"));
            xls.SetCellValue(286, 6, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(286, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(286, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(286, 8, xls.AddFormat(fmt));
            xls.SetCellValue(286, 8, new TFormula("=(    1/ IF(E286<>1,VLOOKUP(E286,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) )  "
            + "    *      IF(F286<>1,VLOOKUP(F286,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)  "
            + "    * IF(G286<>1,VLOOKUP(G286,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(286, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(286, 9, xls.AddFormat(fmt));
            xls.SetCellValue(286, 9, new TFormula("=C286*H286"));

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
            xls.SetCellValue(287, 3, 0);
            xls.SetCellValue(287, 5, new TFormula("='Gral Conf. Summary'!$H$33"));
            xls.SetCellValue(287, 6, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(287, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(287, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(287, 8, xls.AddFormat(fmt));
            xls.SetCellValue(287, 8, new TFormula("=(    1/ IF(E287<>1,VLOOKUP(E287,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) )  "
            + "    *      IF(F287<>1,VLOOKUP(F287,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)  "
            + "    * IF(G287<>1,VLOOKUP(G287,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(287, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(287, 9, xls.AddFormat(fmt));
            xls.SetCellValue(287, 9, new TFormula("=C287*H287"));

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

            fmt = xls.GetCellVisibleFormatDef(288, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(288, 9, xls.AddFormat(fmt));

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
            xls.SetCellValue(289, 3, 439.413333333333);
            xls.SetCellValue(289, 5, new TFormula("='Gral Conf. Summary'!$H$33"));
            xls.SetCellValue(289, 6, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(289, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(289, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(289, 8, xls.AddFormat(fmt));
            xls.SetCellValue(289, 8, new TFormula("=(    1/ IF(E289<>1,VLOOKUP(E289,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) )  "
            + "    *      IF(F289<>1,VLOOKUP(F289,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)  "
            + "    * IF(G289<>1,VLOOKUP(G289,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(289, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(289, 9, xls.AddFormat(fmt));
            xls.SetCellValue(289, 9, new TFormula("=C289*H289"));

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
            xls.SetCellValue(290, 3, 1195.8);
            xls.SetCellValue(290, 5, new TFormula("='Gral Conf. Summary'!$H$33"));
            xls.SetCellValue(290, 6, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(290, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(290, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(290, 8, xls.AddFormat(fmt));
            xls.SetCellValue(290, 8, new TFormula("=(    1/ IF(E290<>1,VLOOKUP(E290,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) )  "
            + "    *      IF(F290<>1,VLOOKUP(F290,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)  "
            + "    * IF(G290<>1,VLOOKUP(G290,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(290, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(290, 9, xls.AddFormat(fmt));
            xls.SetCellValue(290, 9, new TFormula("=C290*H290"));

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
            xls.SetCellValue(291, 3, new TFormula("='Inputs 1.0 Conv. new values'!$M$17"));
            xls.SetCellValue(291, 5, new TFormula("='Gral Conf. Summary'!$H$33"));
            xls.SetCellValue(291, 6, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(291, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(291, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(291, 8, xls.AddFormat(fmt));
            xls.SetCellValue(291, 8, new TFormula("=(    1/ IF(E291<>1,VLOOKUP(E291,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) )  "
            + "    *      IF(F291<>1,VLOOKUP(F291,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)  "
            + "    * IF(G291<>1,VLOOKUP(G291,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(291, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(291, 9, xls.AddFormat(fmt));
            xls.SetCellValue(291, 9, new TFormula("=C291*H291"));

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
            xls.SetCellValue(292, 3, 751.2);
            xls.SetCellValue(292, 5, new TFormula("='Gral Conf. Summary'!$H$33"));
            xls.SetCellValue(292, 6, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(292, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(292, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(292, 8, xls.AddFormat(fmt));
            xls.SetCellValue(292, 8, new TFormula("=(    1/ IF(E292<>1,VLOOKUP(E292,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) )  "
            + "    *      IF(F292<>1,VLOOKUP(F292,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)  "
            + "    * IF(G292<>1,VLOOKUP(G292,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(292, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(292, 9, xls.AddFormat(fmt));
            xls.SetCellValue(292, 9, new TFormula("=C292*H292"));

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
            xls.SetCellValue(293, 3, 0);
            xls.SetCellValue(293, 5, new TFormula("='Gral Conf. Summary'!$H$33"));
            xls.SetCellValue(293, 6, new TFormula("=+'Gral Conf. Summary'!$H$23"));
            xls.SetCellValue(293, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(293, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(293, 8, xls.AddFormat(fmt));
            xls.SetCellValue(293, 8, new TFormula("=(    1/ IF(E293<>1,VLOOKUP(E293,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1) )  "
            + "    *      IF(F293<>1,VLOOKUP(F293,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)  "
            + "    * IF(G293<>1,VLOOKUP(G293,'Gral Conf. Summary'!$K$10:$L$50,2,FALSE),1)"));

            fmt = xls.GetCellVisibleFormatDef(293, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(293, 9, xls.AddFormat(fmt));
            xls.SetCellValue(293, 9, new TFormula("=C293*H293"));

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
            xls.SetComment(10, 5, new TRichString("Juan Hernandez:\nResume all metric used in each question.\n Ej: How many pesos expend"
            + " per hectare?\n\nIn this case the option is:\npesos  hectare 1\n\nTrhere is space"
            + " for 3 simulatanous metrics, if only one, keep the other two as 1 and 1\n\nEj: How"
            + " many quintales?\nquintales 1 1 \n\n\n", Runs, xls));

            //You probably don't need to call the lines below. This code is needed only if you want to change the comment box properties like color or default location
            TCommentProperties CommentProps = TCommentProperties.CreateStandard(10, 5, xls);
            CommentProps.Anchor = new TClientAnchor(TFlxAnchorType.DontMoveAndDontResize, 9, 121, 1, 0, 24, 36, 1, 0);

            //Excel by doesn't autofit the comment box so it can hold all text.
            //There is an option in TCommentProperties, but if you use it Excel will show the text in a single line.
            //To have FlexCel autofit the comment for you, you can do it with the following code:

            //    CommentProps.Anchor = xls.AutofitComment(new TRichString("Juan Hernandez:\nResume all metric used in each question.\n Ej: How many pesos expend"
            //    + " per hectare?\n\nIn this case the option is:\npesos  hectare 1\n\nTrhere is space"
            //    + " for 3 simulatanous metrics, if only one, keep the other two as 1 and 1\n\nEj: How"
            //    + " many quintales?\nquintales 1 1 \n\n\n", Runs, xls), 1.5, true, 1.1, 0, CommentProps.Anchor);

            xls.SetCommentProperties(10, 5, CommentProps);

            //Cell selection and scroll position.
            xls.SelectCell(13, 5, false);
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
            //    xls.DocumentProperties.SetStandardProperty(TPropertyId.CreateTimeDate, "2015-01-08T03:31:31Z");

        }

    }
}
