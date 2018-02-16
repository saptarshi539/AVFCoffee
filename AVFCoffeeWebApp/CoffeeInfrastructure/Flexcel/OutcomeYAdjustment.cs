using CoffeeCore.Interfaces;
using FlexCel.Core;
using FlexCel.XlsAdapter;
using System;
using System.Collections.Generic;
using System.IO;


namespace CoffeeInfrastructure.Flexcel
{
    public class OutcomeYAdjustment
    {
        public void Outcome_Y_Adjustment(ExcelFile xls)
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

            xls.ActiveSheet = 5;    //Set the sheet we are working in.

            //Global Workbook Options
            xls.OptionsAutoCompressPictures = false;
            xls.OptionsCheckCompatibility = false;
            xls.OptionsMultithreadRecalc = 0;

            //Sheet Options
            xls.SheetName = "Outcome_Y_Adjustment";

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
            xls.DefaultColWidth = 2784;

            xls.SetColWidth(1, 2, 2784);    //(10.13 + 0.75) * 256

            xls.SetColWidth(3, 3, 5664);    //(21.38 + 0.75) * 256

            xls.SetColWidth(4, 4, 6496);    //(24.63 + 0.75) * 256

            xls.SetColWidth(5, 5, 3712);    //(13.75 + 0.75) * 256

            xls.SetColWidth(6, 6, 3552);    //(13.13 + 0.75) * 256

            xls.SetColWidth(7, 7, 3744);    //(13.88 + 0.75) * 256

            xls.SetColWidth(8, 8, 4320);    //(16.13 + 0.75) * 256

            xls.SetColWidth(9, 9, 5408);    //(20.38 + 0.75) * 256

            xls.SetColWidth(10, 15, 2784);    //(10.13 + 0.75) * 256

            xls.SetColWidth(16, 16, 3744);    //(13.88 + 0.75) * 256

            xls.SetColWidth(17, 16384, 2784);    //(10.13 + 0.75) * 256
            xls.DefaultRowHeight = 315;

            xls.SetRowHeight(2, 960);    //48.00 * 20

            //Set the cell values
            TFlxFormat fmt;
            fmt = xls.GetCellVisibleFormatDef(2, 3);
            fmt.WrapText = true;
            xls.SetCellFormat(2, 3, xls.AddFormat(fmt));
            xls.SetCellValue(2, 3, "Productivity original scale (Quintales/ht)");
            xls.SetCellValue(2, 4, new TFormula("='Inputs 1.0_metric_currency'!$D$15"));
            xls.SetCellValue(3, 3, "Productivity Pounds/ht");
            xls.SetCellValue(3, 4, new TFormula("=D2*Conversiones!C14"));
            xls.SetCellValue(5, 7, "x");
            xls.SetCellValue(5, 8, "y");
            xls.SetCellValue(5, 9, "y pesos");
            xls.SetCellValue(6, 4, "a");
            xls.SetCellValue(6, 5, "b");
            xls.SetCellValue(6, 6, "c");
            xls.SetCellValue(6, 7, "Pounds/ht");
            xls.SetCellValue(6, 8, "Cost (US/ht)");
            xls.SetCellValue(6, 9, "Cost (Pesos/ht)");
            xls.SetCellValue(7, 3, "Variable");

            fmt = xls.GetCellVisibleFormatDef(7, 4);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.Format = "0.00E+00";
            xls.SetCellFormat(7, 4, xls.AddFormat(fmt));
            xls.SetCellValue(7, 4, 0.000424373358854431);

            fmt = xls.GetCellVisibleFormatDef(7, 5);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            xls.SetCellFormat(7, 5, xls.AddFormat(fmt));
            xls.SetCellValue(7, 5, -1.33320389343239);

            fmt = xls.GetCellVisibleFormatDef(7, 6);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            xls.SetCellFormat(7, 6, xls.AddFormat(fmt));
            xls.SetCellValue(7, 6, 2199.65348186419);

            fmt = xls.GetCellVisibleFormatDef(7, 7);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(7, 7, xls.AddFormat(fmt));
            xls.SetCellValue(7, 7, new TFormula("=$D$3"));
            xls.SetCellValue(7, 8, new TFormula("=$D7*(G7^2)+($E7*G7)+$F7"));
            xls.SetCellValue(7, 9, new TFormula("=H7*Conversiones!$F$24"));
            xls.SetCellValue(8, 3, "Fixed");

            fmt = xls.GetCellVisibleFormatDef(8, 4);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            xls.SetCellFormat(8, 4, xls.AddFormat(fmt));
            xls.SetCellValue(8, 4, 0.000434909511174162);

            fmt = xls.GetCellVisibleFormatDef(8, 5);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            xls.SetCellFormat(8, 5, xls.AddFormat(fmt));
            xls.SetCellValue(8, 5, -1.31505293454055);

            fmt = xls.GetCellVisibleFormatDef(8, 6);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            xls.SetCellFormat(8, 6, xls.AddFormat(fmt));
            xls.SetCellValue(8, 6, 2196.17801790681);

            fmt = xls.GetCellVisibleFormatDef(8, 7);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(8, 7, xls.AddFormat(fmt));
            xls.SetCellValue(8, 7, new TFormula("=$D$3"));
            xls.SetCellValue(8, 8, new TFormula("=$D8*(G8^2)+($E8*G8)+$F8"));
            xls.SetCellValue(8, 9, new TFormula("=H8*Conversiones!$F$24"));
            xls.SetCellValue(9, 3, "Depreciation");

            fmt = xls.GetCellVisibleFormatDef(9, 4);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            xls.SetCellFormat(9, 4, xls.AddFormat(fmt));
            xls.SetCellValue(9, 4, 0.000387864040252227);

            fmt = xls.GetCellVisibleFormatDef(9, 5);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            xls.SetCellFormat(9, 5, xls.AddFormat(fmt));
            xls.SetCellValue(9, 5, -0.980675788815258);

            fmt = xls.GetCellVisibleFormatDef(9, 6);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            xls.SetCellFormat(9, 6, xls.AddFormat(fmt));
            xls.SetCellValue(9, 6, 2356.26888204398);

            fmt = xls.GetCellVisibleFormatDef(9, 7);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(9, 7, xls.AddFormat(fmt));
            xls.SetCellValue(9, 7, new TFormula("=$D$3"));
            xls.SetCellValue(9, 8, new TFormula("=$D9*(G9^2)+($E9*G9)+$F9"));
            xls.SetCellValue(9, 9, new TFormula("=H9*Conversiones!$F$24"));
            xls.SetCellValue(10, 3, "Total");

            fmt = xls.GetCellVisibleFormatDef(10, 4);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            xls.SetCellFormat(10, 4, xls.AddFormat(fmt));
            xls.SetCellValue(10, 4, 0.000545755328389956);

            fmt = xls.GetCellVisibleFormatDef(10, 5);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            xls.SetCellFormat(10, 5, xls.AddFormat(fmt));
            xls.SetCellValue(10, 5, -1.70794924924928);

            fmt = xls.GetCellVisibleFormatDef(10, 6);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            xls.SetCellFormat(10, 6, xls.AddFormat(fmt));
            xls.SetCellValue(10, 6, 3557.11221066245);

            fmt = xls.GetCellVisibleFormatDef(10, 7);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(10, 7, xls.AddFormat(fmt));
            xls.SetCellValue(10, 7, new TFormula("=$D$3"));
            xls.SetCellValue(10, 8, new TFormula("=$D10*(G10^2)+($E10*G10)+$F10"));
            xls.SetCellValue(10, 9, new TFormula("=H10*Conversiones!$F$24"));

            //Cell selection and scroll position.
            xls.SelectCell(32, 8, false);
            xls.ScrollWindow(13, 1);

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
