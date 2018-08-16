using FlexCel.Core;
using System;
using System.Collections.Generic;
using System.Text;

namespace CoffeeInfrastructure.Flexcel
{
    public class Metrics_English
    {
        public void MetricsEnglish(ExcelFile xls)
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

            xls.ActiveSheet = 3;    //Set the sheet we are working in.

            //Global Workbook Options
            xls.OptionsAutoCompressPictures = false;
            xls.OptionsMultithreadRecalc = 0;

            //Sheet Options
            xls.SheetName = "Metrics English";

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

            TFlxFormat ColFmt;
            ColFmt = xls.GetFormat(xls.GetColFormat(4));
            ColFmt.HAlignment = THFlxAlignment.center;
            xls.SetColFormat(4, 4, xls.AddFormat(ColFmt));

            xls.SetColWidth(5, 16384, 2784);    //(10.13 + 0.75) * 256
            xls.DefaultRowHeight = 315;

            xls.SetRowHeight(2, 330);    //16.50 * 20
            xls.SetRowHeight(3, 375);    //18.75 * 20
            xls.SetRowHeight(10, 375);    //18.75 * 20
            xls.SetRowHeight(12, 375);    //18.75 * 20
            xls.SetRowHeight(47, 330);    //16.50 * 20

            //Merged Cells
            xls.MergeCells(3, 2, 3, 6);
            xls.MergeCells(5, 2, 5, 6);

            //Set the cell values
            TFlxFormat fmt;
            fmt = xls.GetCellVisibleFormatDef(2, 2);
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            xls.SetCellFormat(2, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(2, 3);
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            xls.SetCellFormat(2, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(2, 4);
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(2, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(2, 5);
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            xls.SetCellFormat(2, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(2, 6);
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            xls.SetCellFormat(2, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(3, 1);
            fmt.Borders.Right.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            xls.SetCellFormat(3, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(3, 2);
            fmt.Font.Size20 = 280;
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0x00, 0x80, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(3, 2, xls.AddFormat(fmt));
            xls.SetCellValue(3, 2, "METRICS");

            fmt = xls.GetCellVisibleFormatDef(3, 3);
            fmt.Font.Size20 = 280;
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0x00, 0x80, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(3, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(3, 4);
            fmt.Font.Size20 = 280;
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0x00, 0x80, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(3, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(3, 5);
            fmt.Font.Size20 = 280;
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0x00, 0x80, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(3, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(3, 6);
            fmt.Font.Size20 = 280;
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0x00, 0x80, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(3, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(3, 7);
            fmt.Font.Size20 = 280;
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(3, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(3, 8);
            fmt.Font.Size20 = 280;
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(3, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(3, 9);
            fmt.Font.Size20 = 280;
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(3, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(3, 10);
            fmt.Font.Size20 = 280;
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(3, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(3, 11);
            fmt.Font.Size20 = 280;
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(3, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(3, 12);
            fmt.Font.Size20 = 280;
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(3, 12, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(3, 13);
            fmt.Font.Size20 = 280;
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(3, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(4, 1);
            fmt.Borders.Right.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            xls.SetCellFormat(4, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(4, 2);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(4, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(4, 3);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(4, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(4, 4);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(4, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(4, 5);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(4, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(4, 6);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(4, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(4, 7);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(4, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(5, 1);
            fmt.Borders.Right.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            xls.SetCellFormat(5, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(5, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(5, 2, xls.AddFormat(fmt));
            xls.SetCellValue(5, 2, "Metric and meassure system definition");

            fmt = xls.GetCellVisibleFormatDef(5, 3);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(5, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(5, 4);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(5, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(5, 5);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(5, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(5, 6);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(5, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(5, 7);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(5, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(6, 1);
            fmt.Borders.Right.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            xls.SetCellFormat(6, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(6, 2);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(6, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(6, 3);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(6, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(6, 4);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(6, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(6, 5);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(6, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(6, 6);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(6, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(6, 7);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(6, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(7, 1);
            fmt.Borders.Right.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            xls.SetCellFormat(7, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(7, 2);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(7, 2, xls.AddFormat(fmt));
            xls.SetCellValue(7, 2, "Please answer the following questions and select one option in each question");

            fmt = xls.GetCellVisibleFormatDef(7, 3);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(7, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(7, 4);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(7, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(7, 5);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(7, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(7, 6);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(7, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(7, 7);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(7, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(8, 1);
            fmt.Borders.Right.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            xls.SetCellFormat(8, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(8, 2);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(8, 2, xls.AddFormat(fmt));
            xls.SetCellValue(8, 2, "(Write 1 for the selected option)");

            fmt = xls.GetCellVisibleFormatDef(8, 3);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(8, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(8, 4);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(8, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(8, 5);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(8, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(8, 6);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(8, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(8, 7);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(8, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(9, 1);
            fmt.Borders.Right.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            xls.SetCellFormat(9, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(9, 2);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(9, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(9, 3);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(9, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(9, 4);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(9, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(9, 5);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(9, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(9, 6);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(9, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(9, 7);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(9, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(10, 1);
            fmt.Borders.Right.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            xls.SetCellFormat(10, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(10, 2);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(10, 2, xls.AddFormat(fmt));
            xls.SetCellValue(10, 2, "The most commom meassure for coffee parchment is:");

            fmt = xls.GetCellVisibleFormatDef(10, 3);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(10, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(10, 4);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(10, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(10, 5);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(10, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(10, 6);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(10, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(10, 7);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(10, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(10, 8);
            fmt.Font.Size20 = 280;
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Family = 0;
            xls.SetCellFormat(10, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(10, 13);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(10, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(10, 15);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(10, 15, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(11, 1);
            fmt.Borders.Right.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            xls.SetCellFormat(11, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(11, 2);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(11, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(11, 3);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.Font.Style = TFlxFontStyles.Italic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(11, 3, xls.AddFormat(fmt));
            xls.SetCellValue(11, 3, "Kilograms");

            fmt = xls.GetCellVisibleFormatDef(11, 4);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(11, 4, xls.AddFormat(fmt));
            xls.SetCellValue(11, 4, 1);

            fmt = xls.GetCellVisibleFormatDef(11, 5);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(11, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(11, 6);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(11, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(11, 7);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(11, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(12, 1);
            fmt.Borders.Right.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            xls.SetCellFormat(12, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(12, 2);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(12, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(12, 3);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.Font.Style = TFlxFontStyles.Italic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(12, 3, xls.AddFormat(fmt));
            xls.SetCellValue(12, 3, "Pounds");

            fmt = xls.GetCellVisibleFormatDef(12, 4);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(12, 4, xls.AddFormat(fmt));
            xls.SetCellValue(12, 4, 0);

            fmt = xls.GetCellVisibleFormatDef(12, 5);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(12, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(12, 6);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(12, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(12, 7);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(12, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(12, 8);
            fmt.Font.Size20 = 280;
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Family = 0;
            xls.SetCellFormat(12, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(13, 1);
            fmt.Borders.Right.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            xls.SetCellFormat(13, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(13, 2);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(13, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(13, 3);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.Font.Style = TFlxFontStyles.Italic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(13, 3, xls.AddFormat(fmt));
            xls.SetCellValue(13, 3, "Quintales");

            fmt = xls.GetCellVisibleFormatDef(13, 4);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(13, 4, xls.AddFormat(fmt));
            xls.SetCellValue(13, 4, 0);

            fmt = xls.GetCellVisibleFormatDef(13, 5);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(13, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(13, 6);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(13, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(13, 7);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(13, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(14, 1);
            fmt.Borders.Right.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            xls.SetCellFormat(14, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(14, 2);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(14, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(14, 3);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.Font.Style = TFlxFontStyles.Italic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(14, 3, xls.AddFormat(fmt));
            xls.SetCellValue(14, 3, "Arrobas");

            fmt = xls.GetCellVisibleFormatDef(14, 4);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(14, 4, xls.AddFormat(fmt));
            xls.SetCellValue(14, 4, 0);

            fmt = xls.GetCellVisibleFormatDef(14, 5);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(14, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(14, 6);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(14, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(14, 7);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(14, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(15, 1);
            fmt.Borders.Right.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            xls.SetCellFormat(15, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(15, 2);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(15, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(15, 3);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.Font.Style = TFlxFontStyles.Italic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(15, 3, xls.AddFormat(fmt));
            xls.SetCellValue(15, 3, "Cargas");

            fmt = xls.GetCellVisibleFormatDef(15, 4);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(15, 4, xls.AddFormat(fmt));
            xls.SetCellValue(15, 4, 0);

            fmt = xls.GetCellVisibleFormatDef(15, 5);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(15, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(15, 6);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(15, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(15, 7);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(15, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(16, 1);
            fmt.Borders.Right.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            xls.SetCellFormat(16, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(16, 2);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(16, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(16, 3);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(16, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(16, 4);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(16, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(16, 5);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(16, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(16, 6);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(16, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(16, 7);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(16, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(17, 1);
            fmt.Borders.Right.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            xls.SetCellFormat(17, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(17, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(17, 2, xls.AddFormat(fmt));
            xls.SetCellValue(17, 2, "The most commom meassure of length is:");

            fmt = xls.GetCellVisibleFormatDef(17, 3);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(17, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(17, 4);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(17, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(17, 5);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(17, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(17, 6);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(17, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(17, 7);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(17, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(17, 13);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(17, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(18, 1);
            fmt.Borders.Right.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            xls.SetCellFormat(18, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(18, 2);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(18, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(18, 3);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(18, 3, xls.AddFormat(fmt));
            xls.SetCellValue(18, 3, "Meters");

            fmt = xls.GetCellVisibleFormatDef(18, 4);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(18, 4, xls.AddFormat(fmt));
            xls.SetCellValue(18, 4, 1);

            fmt = xls.GetCellVisibleFormatDef(18, 5);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(18, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(18, 6);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(18, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(18, 7);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(18, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(19, 1);
            fmt.Borders.Right.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            xls.SetCellFormat(19, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(19, 2);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(19, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(19, 3);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(19, 3, xls.AddFormat(fmt));
            xls.SetCellValue(19, 3, "Feet");

            fmt = xls.GetCellVisibleFormatDef(19, 4);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(19, 4, xls.AddFormat(fmt));
            xls.SetCellValue(19, 4, 0);

            fmt = xls.GetCellVisibleFormatDef(19, 5);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(19, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(19, 6);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(19, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(19, 7);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(19, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(20, 1);
            fmt.Borders.Right.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            xls.SetCellFormat(20, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(20, 2);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(20, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(20, 3);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(20, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(20, 4);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(20, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(20, 5);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(20, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(20, 6);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(20, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(20, 7);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(20, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(21, 1);
            fmt.Borders.Right.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            xls.SetCellFormat(21, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(21, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(21, 2, xls.AddFormat(fmt));
            xls.SetCellValue(21, 2, "The most commom meassure for the farm area is:");

            fmt = xls.GetCellVisibleFormatDef(21, 3);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(21, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(21, 4);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(21, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(21, 5);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(21, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(21, 6);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(21, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(21, 7);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(21, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(21, 13);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(21, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(22, 1);
            fmt.Borders.Right.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            xls.SetCellFormat(22, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(22, 2);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(22, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(22, 3);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(22, 3, xls.AddFormat(fmt));
            xls.SetCellValue(22, 3, "Hectares");

            fmt = xls.GetCellVisibleFormatDef(22, 4);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(22, 4, xls.AddFormat(fmt));
            xls.SetCellValue(22, 4, 1);

            fmt = xls.GetCellVisibleFormatDef(22, 5);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(22, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(22, 6);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(22, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(22, 7);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(22, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(23, 1);
            fmt.Borders.Right.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            xls.SetCellFormat(23, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(23, 2);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(23, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(23, 3);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(23, 3, xls.AddFormat(fmt));
            xls.SetCellValue(23, 3, "Manzanas");

            fmt = xls.GetCellVisibleFormatDef(23, 4);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(23, 4, xls.AddFormat(fmt));
            xls.SetCellValue(23, 4, 0);

            fmt = xls.GetCellVisibleFormatDef(23, 5);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(23, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(23, 6);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(23, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(23, 7);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(23, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(24, 1);
            fmt.Borders.Right.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            xls.SetCellFormat(24, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(24, 2);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(24, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(24, 3);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(24, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(24, 4);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(24, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(24, 5);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(24, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(24, 6);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(24, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(24, 7);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(24, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(25, 1);
            fmt.Borders.Right.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            xls.SetCellFormat(25, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(25, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(25, 2, xls.AddFormat(fmt));
            xls.SetCellValue(25, 2, "The most commom meassure for inputs weight (Ex: fertilizers is):");

            fmt = xls.GetCellVisibleFormatDef(25, 3);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(25, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(25, 4);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(25, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(25, 5);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(25, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(25, 6);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(25, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(25, 7);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(25, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(25, 13);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(25, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(26, 1);
            fmt.Borders.Right.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            xls.SetCellFormat(26, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(26, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(26, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(26, 3);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.Font.Style = TFlxFontStyles.Italic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(26, 3, xls.AddFormat(fmt));
            xls.SetCellValue(26, 3, "Kilograms");

            fmt = xls.GetCellVisibleFormatDef(26, 4);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(26, 4, xls.AddFormat(fmt));
            xls.SetCellValue(26, 4, 1);

            fmt = xls.GetCellVisibleFormatDef(26, 5);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(26, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(26, 6);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(26, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(26, 7);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(26, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(27, 1);
            fmt.Borders.Right.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            xls.SetCellFormat(27, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(27, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(27, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(27, 3);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(27, 3, xls.AddFormat(fmt));
            xls.SetCellValue(27, 3, "Pounds");

            fmt = xls.GetCellVisibleFormatDef(27, 4);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(27, 4, xls.AddFormat(fmt));
            xls.SetCellValue(27, 4, 0);

            fmt = xls.GetCellVisibleFormatDef(27, 5);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(27, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(27, 6);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(27, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(27, 7);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(27, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(28, 1);
            fmt.Borders.Right.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            xls.SetCellFormat(28, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(28, 2);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(28, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(28, 3);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(28, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(28, 4);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(28, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(28, 5);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(28, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(28, 6);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(28, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(28, 7);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(28, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(29, 1);
            fmt.Borders.Right.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            xls.SetCellFormat(29, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(29, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(29, 2, xls.AddFormat(fmt));
            xls.SetCellValue(29, 2, "The most commom meassure for capacity:");

            fmt = xls.GetCellVisibleFormatDef(29, 3);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(29, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(29, 4);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(29, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(29, 5);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(29, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(29, 6);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(29, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(29, 7);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(29, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(29, 13);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(29, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(30, 1);
            fmt.Borders.Right.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            xls.SetCellFormat(30, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(30, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(30, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(30, 3);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(30, 3, xls.AddFormat(fmt));
            xls.SetCellValue(30, 3, "Liters");

            fmt = xls.GetCellVisibleFormatDef(30, 4);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(30, 4, xls.AddFormat(fmt));
            xls.SetCellValue(30, 4, 1);

            fmt = xls.GetCellVisibleFormatDef(30, 5);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(30, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(30, 6);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(30, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(30, 7);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(30, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(31, 1);
            fmt.Borders.Right.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            xls.SetCellFormat(31, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(31, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(31, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(31, 3);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(31, 3, xls.AddFormat(fmt));
            xls.SetCellValue(31, 3, "Gallons");

            fmt = xls.GetCellVisibleFormatDef(31, 4);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(31, 4, xls.AddFormat(fmt));
            xls.SetCellValue(31, 4, 0);

            fmt = xls.GetCellVisibleFormatDef(31, 5);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(31, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(31, 6);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(31, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(31, 7);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(31, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(32, 1);
            fmt.Borders.Right.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            xls.SetCellFormat(32, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(32, 2);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(32, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(32, 3);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(32, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(32, 4);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(32, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(32, 5);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(32, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(32, 6);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(32, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(32, 7);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(32, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(33, 1);
            fmt.Borders.Right.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            xls.SetCellFormat(33, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(33, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(33, 2, xls.AddFormat(fmt));
            xls.SetCellValue(33, 2, "Country currency");

            fmt = xls.GetCellVisibleFormatDef(33, 3);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(33, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(33, 4);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(33, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(33, 5);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(33, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(33, 6);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(33, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(33, 7);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(33, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(33, 13);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(33, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(34, 1);
            fmt.Borders.Right.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            xls.SetCellFormat(34, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(34, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(34, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(34, 3);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(34, 3, xls.AddFormat(fmt));
            xls.SetCellValue(34, 3, "Bolivian boliviano");

            fmt = xls.GetCellVisibleFormatDef(34, 4);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(34, 4, xls.AddFormat(fmt));
            xls.SetCellValue(34, 4, 0);

            fmt = xls.GetCellVisibleFormatDef(34, 5);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(34, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(34, 6);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(34, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(34, 7);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(34, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(34, 13);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(34, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(35, 1);
            fmt.Borders.Right.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            xls.SetCellFormat(35, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(35, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(35, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(35, 3);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(35, 3, xls.AddFormat(fmt));
            xls.SetCellValue(35, 3, "Brazilian real");

            fmt = xls.GetCellVisibleFormatDef(35, 4);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(35, 4, xls.AddFormat(fmt));
            xls.SetCellValue(35, 4, 0);

            fmt = xls.GetCellVisibleFormatDef(35, 5);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(35, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(35, 6);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(35, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(35, 7);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(35, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(35, 13);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(35, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(36, 1);
            fmt.Borders.Right.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            xls.SetCellFormat(36, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(36, 2);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(36, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(36, 3);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(36, 3, xls.AddFormat(fmt));
            xls.SetCellValue(36, 3, "Colombian peso");

            fmt = xls.GetCellVisibleFormatDef(36, 4);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(36, 4, xls.AddFormat(fmt));
            xls.SetCellValue(36, 4, 0);

            fmt = xls.GetCellVisibleFormatDef(36, 5);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(36, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(36, 6);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(36, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(36, 7);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(36, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(37, 1);
            fmt.Borders.Right.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            xls.SetCellFormat(37, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(37, 2);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(37, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(37, 3);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(37, 3, xls.AddFormat(fmt));
            xls.SetCellValue(37, 3, "Costa rican colon");

            fmt = xls.GetCellVisibleFormatDef(37, 4);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(37, 4, xls.AddFormat(fmt));
            xls.SetCellValue(37, 4, 0);

            fmt = xls.GetCellVisibleFormatDef(37, 5);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(37, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(37, 6);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(37, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(37, 7);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(37, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(38, 1);
            fmt.Borders.Right.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            xls.SetCellFormat(38, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(38, 2);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(38, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(38, 3);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(38, 3, xls.AddFormat(fmt));
            xls.SetCellValue(38, 3, "Cuban peso");

            fmt = xls.GetCellVisibleFormatDef(38, 4);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(38, 4, xls.AddFormat(fmt));
            xls.SetCellValue(38, 4, 0);

            fmt = xls.GetCellVisibleFormatDef(38, 5);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(38, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(38, 6);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(38, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(38, 7);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(38, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(39, 1);
            fmt.Borders.Right.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            xls.SetCellFormat(39, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(39, 2);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(39, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(39, 3);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(39, 3, xls.AddFormat(fmt));
            xls.SetCellValue(39, 3, "Guatemalan quetzal");

            fmt = xls.GetCellVisibleFormatDef(39, 4);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(39, 4, xls.AddFormat(fmt));
            xls.SetCellValue(39, 4, 0);

            fmt = xls.GetCellVisibleFormatDef(39, 5);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(39, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(39, 6);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(39, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(39, 7);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(39, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(40, 1);
            fmt.Borders.Right.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            xls.SetCellFormat(40, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(40, 2);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(40, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(40, 3);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(40, 3, xls.AddFormat(fmt));
            xls.SetCellValue(40, 3, "Haitian gourde");

            fmt = xls.GetCellVisibleFormatDef(40, 4);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(40, 4, xls.AddFormat(fmt));
            xls.SetCellValue(40, 4, 0);

            fmt = xls.GetCellVisibleFormatDef(40, 5);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(40, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(40, 6);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(40, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(40, 7);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(40, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(41, 1);
            fmt.Borders.Right.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            xls.SetCellFormat(41, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(41, 2);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(41, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(41, 3);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(41, 3, xls.AddFormat(fmt));
            xls.SetCellValue(41, 3, "Honduran lempira");

            fmt = xls.GetCellVisibleFormatDef(41, 4);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(41, 4, xls.AddFormat(fmt));
            xls.SetCellValue(41, 4, 0);

            fmt = xls.GetCellVisibleFormatDef(41, 5);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(41, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(41, 6);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(41, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(41, 7);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(41, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(42, 1);
            fmt.Borders.Right.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            xls.SetCellFormat(42, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(42, 2);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(42, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(42, 3);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(42, 3, xls.AddFormat(fmt));
            xls.SetCellValue(42, 3, "Jamaican dollar");

            fmt = xls.GetCellVisibleFormatDef(42, 4);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(42, 4, xls.AddFormat(fmt));
            xls.SetCellValue(42, 4, 0);

            fmt = xls.GetCellVisibleFormatDef(42, 5);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(42, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(42, 6);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(42, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(42, 7);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(42, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(43, 1);
            fmt.Borders.Right.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            xls.SetCellFormat(43, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(43, 2);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(43, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(43, 3);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(43, 3, xls.AddFormat(fmt));
            xls.SetCellValue(43, 3, "Mexican peso");

            fmt = xls.GetCellVisibleFormatDef(43, 4);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(43, 4, xls.AddFormat(fmt));
            xls.SetCellValue(43, 4, 1);

            fmt = xls.GetCellVisibleFormatDef(43, 5);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(43, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(43, 6);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(43, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(43, 7);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(43, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(44, 1);
            fmt.Borders.Right.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            xls.SetCellFormat(44, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(44, 2);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(44, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(44, 3);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(44, 3, xls.AddFormat(fmt));
            xls.SetCellValue(44, 3, "Nicaraguan cordoba");

            fmt = xls.GetCellVisibleFormatDef(44, 4);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(44, 4, xls.AddFormat(fmt));
            xls.SetCellValue(44, 4, 0);

            fmt = xls.GetCellVisibleFormatDef(44, 5);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(44, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(44, 6);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(44, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(44, 7);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(44, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(45, 1);
            fmt.Borders.Right.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            xls.SetCellFormat(45, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(45, 2);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(45, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(45, 3);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(45, 3, xls.AddFormat(fmt));
            xls.SetCellValue(45, 3, "Peruvian sol");

            fmt = xls.GetCellVisibleFormatDef(45, 4);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(45, 4, xls.AddFormat(fmt));
            xls.SetCellValue(45, 4, 0);

            fmt = xls.GetCellVisibleFormatDef(45, 5);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(45, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(45, 6);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(45, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(45, 7);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(45, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(46, 1);
            fmt.Borders.Right.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            xls.SetCellFormat(46, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(46, 2);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(46, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(46, 3);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(46, 3, xls.AddFormat(fmt));
            xls.SetCellValue(46, 3, "US dollar");

            fmt = xls.GetCellVisibleFormatDef(46, 4);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(46, 4, xls.AddFormat(fmt));
            xls.SetCellValue(46, 4, 0);

            fmt = xls.GetCellVisibleFormatDef(46, 5);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(46, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(46, 6);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(46, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(46, 7);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(46, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(47, 1);
            fmt.Borders.Right.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            xls.SetCellFormat(47, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(47, 2);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(47, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(47, 3);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(47, 3, xls.AddFormat(fmt));
            xls.SetCellValue(47, 3, "Venezuelan bolivar");

            fmt = xls.GetCellVisibleFormatDef(47, 4);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(47, 4, xls.AddFormat(fmt));
            xls.SetCellValue(47, 4, 0);

            fmt = xls.GetCellVisibleFormatDef(47, 5);
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(47, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(47, 6);
            fmt.Borders.Right.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(47, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(47, 7);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            xls.SetCellFormat(47, 7, xls.AddFormat(fmt));

            //Cell selection and scroll position.
            xls.SelectCell(34, 4, false);
            xls.ScrollWindow(14, 1);

            //Standard Document Properties - Most are only for xlsx files. In xls files FlexCel will only change the Creation Date and Modified Date.
            xls.DocumentProperties.SetStandardProperty(TPropertyId.Author, "Mary Kate");
            xls.Recalc();
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
