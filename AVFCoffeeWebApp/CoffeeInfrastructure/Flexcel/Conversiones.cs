using CoffeeCore.Interfaces;
using FlexCel.Core;
using FlexCel.XlsAdapter;
using System;
using System.Collections.Generic;
using System.IO;

namespace CoffeeInfrastructure.Flexcel
{
    public class Conversiones
    {
        public void conversiones(ExcelFile xls)
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

            xls.ActiveSheet = 18;    //Set the sheet we are working in.

            //Global Workbook Options
            xls.OptionsAutoCompressPictures = false;
            xls.OptionsMultithreadRecalc = 0;

            //Sheet Options
            xls.SheetName = "Conversiones";

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
            xls.DefaultColWidth = 2816;

            xls.SetColWidth(2, 2, 3584);    //(13.25 + 0.75) * 256

            xls.SetColWidth(3, 3, 4448);    //(16.63 + 0.75) * 256

            xls.SetColWidth(4, 4, 3840);    //(14.25 + 0.75) * 256

            xls.SetColWidth(5, 5, 4000);    //(14.88 + 0.75) * 256
            xls.DefaultRowHeight = 315;

            xls.SetRowHeight(3, 375);    //18.75 * 20
            xls.SetRowHeight(4, 375);    //18.75 * 20
            xls.SetRowHeight(33, 630);    //31.50 * 20
            xls.SetRowHeight(34, 630);    //31.50 * 20
            xls.SetRowHeight(36, 630);    //31.50 * 20
            xls.SetRowHeight(51, 630);    //31.50 * 20

            //Merged Cells
            xls.MergeCells(42, 8, 42, 9);
            xls.MergeCells(67, 2, 67, 3);
            xls.MergeCells(72, 2, 72, 3);
            xls.MergeCells(79, 2, 79, 3);
            xls.MergeCells(3, 2, 3, 4);
            xls.MergeCells(42, 2, 42, 3);
            xls.MergeCells(54, 2, 54, 3);

            //Set the cell values
            TFlxFormat fmt;
            fmt = xls.GetCellVisibleFormatDef(3, 2);
            fmt.Font.Size20 = 280;
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(3, 2, xls.AddFormat(fmt));
            xls.SetCellValue(3, 2, "Factores de Conversión");

            fmt = xls.GetCellVisibleFormatDef(3, 3);
            fmt.Font.Size20 = 280;
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(3, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(3, 4);
            fmt.Font.Size20 = 280;
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(3, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(4, 2);
            fmt.Font.Size20 = 280;
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(4, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(4, 3);
            fmt.Font.Size20 = 280;
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(4, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(4, 4);
            fmt.Font.Size20 = 280;
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(4, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(5, 2);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0x33, 0x66, 0xFF);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(5, 2, xls.AddFormat(fmt));
            xls.SetCellValue(5, 2, "Area");

            fmt = xls.GetCellVisibleFormatDef(6, 2);
            fmt.Font.Style = TFlxFontStyles.Italic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent4, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(6, 2, xls.AddFormat(fmt));
            xls.SetCellValue(6, 2, "Hectarea");

            fmt = xls.GetCellVisibleFormatDef(6, 3);
            fmt.Font.Style = TFlxFontStyles.Italic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(6, 3, xls.AddFormat(fmt));
            xls.SetCellValue(6, 3, "Manzana");
            xls.SetCellValue(6, 4, "Metros 2");
            xls.SetCellValue(7, 2, 1);
            xls.SetCellValue(7, 3, 1.4184);
            xls.SetCellValue(7, 4, 10000);
            xls.SetCellValue(8, 11, "Quintal");

            fmt = xls.GetCellVisibleFormatDef(9, 2);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0x33, 0x66, 0xFF);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(9, 2, xls.AddFormat(fmt));
            xls.SetCellValue(9, 2, "Weight");

            fmt = xls.GetCellVisibleFormatDef(10, 2);
            fmt.Font.Style = TFlxFontStyles.Italic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent4, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(10, 2, xls.AddFormat(fmt));
            xls.SetCellValue(10, 2, "Kilo");

            fmt = xls.GetCellVisibleFormatDef(10, 3);
            fmt.Font.Style = TFlxFontStyles.Italic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(10, 3, xls.AddFormat(fmt));
            xls.SetCellValue(10, 3, "Libra");
            xls.SetCellValue(11, 2, 1);

            fmt = xls.GetCellVisibleFormatDef(11, 3);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(11, 3, xls.AddFormat(fmt));
            xls.SetCellValue(11, 3, 2.20462);

            fmt = xls.GetCellVisibleFormatDef(13, 2);
            fmt.Font.Style = TFlxFontStyles.Italic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent4, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(13, 2, xls.AddFormat(fmt));
            xls.SetCellValue(13, 2, "Quintal");

            fmt = xls.GetCellVisibleFormatDef(13, 3);
            fmt.Font.Style = TFlxFontStyles.Italic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(13, 3, xls.AddFormat(fmt));
            xls.SetCellValue(13, 3, "Libra");

            fmt = xls.GetCellVisibleFormatDef(13, 4);
            fmt.Font.Style = TFlxFontStyles.Italic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(13, 4, xls.AddFormat(fmt));
            xls.SetCellValue(13, 4, "Kilo");
            xls.SetCellValue(14, 2, 1);

            fmt = xls.GetCellVisibleFormatDef(14, 3);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(14, 3, xls.AddFormat(fmt));
            xls.SetCellValue(14, 3, 101.4);

            fmt = xls.GetCellVisibleFormatDef(14, 4);
            fmt.Format = "0.00";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(14, 4, xls.AddFormat(fmt));
            xls.SetCellValue(14, 4, new TFormula("=H44"));

            fmt = xls.GetCellVisibleFormatDef(16, 2);
            fmt.Font.Style = TFlxFontStyles.Italic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent4, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(16, 2, xls.AddFormat(fmt));
            xls.SetCellValue(16, 2, "Arroba");

            fmt = xls.GetCellVisibleFormatDef(16, 3);
            fmt.Font.Style = TFlxFontStyles.Italic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(16, 3, xls.AddFormat(fmt));
            xls.SetCellValue(16, 3, "Libra");

            fmt = xls.GetCellVisibleFormatDef(16, 4);
            fmt.Font.Style = TFlxFontStyles.Italic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(16, 4, xls.AddFormat(fmt));
            xls.SetCellValue(16, 4, "Kilo");
            xls.SetCellValue(17, 2, 1);
            xls.SetCellValue(17, 4, 12.5);

            fmt = xls.GetCellVisibleFormatDef(19, 2);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent4, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(19, 2, xls.AddFormat(fmt));
            xls.SetCellValue(19, 2, "Carga");
            xls.SetCellValue(19, 4, "Kilo");
            xls.SetCellValue(19, 5, "Sacos");
            xls.SetCellValue(19, 6, "Arrobas");
            xls.SetCellValue(20, 2, 1);
            xls.SetCellValue(20, 4, 125);
            xls.SetCellValue(20, 5, 2);
            xls.SetCellValue(20, 6, new TFormula("=D20/D17"));

            fmt = xls.GetCellVisibleFormatDef(22, 2);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0x33, 0x66, 0xFF);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(22, 2, xls.AddFormat(fmt));
            xls.SetCellValue(22, 2, "Currency");

            fmt = xls.GetCellVisibleFormatDef(23, 2);
            fmt.Font.Style = TFlxFontStyles.Italic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent4, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(23, 2, xls.AddFormat(fmt));
            xls.SetCellValue(23, 2, "USD");

            fmt = xls.GetCellVisibleFormatDef(23, 3);
            fmt.Font.Style = TFlxFontStyles.Italic;
            xls.SetCellFormat(23, 3, xls.AddFormat(fmt));
            xls.SetCellValue(23, 3, "Lempira (Honduras)");

            fmt = xls.GetCellVisibleFormatDef(23, 4);
            fmt.Font.Style = TFlxFontStyles.Italic;
            xls.SetCellFormat(23, 4, xls.AddFormat(fmt));
            xls.SetCellValue(23, 4, "Nuevo Sol (Peru)");

            fmt = xls.GetCellVisibleFormatDef(23, 5);
            fmt.Font.Style = TFlxFontStyles.Italic;
            xls.SetCellFormat(23, 5, xls.AddFormat(fmt));
            xls.SetCellValue(23, 5, "Pesos  (Colombia)");

            fmt = xls.GetCellVisibleFormatDef(23, 6);
            fmt.Font.Style = TFlxFontStyles.Italic;
            xls.SetCellFormat(23, 6, xls.AddFormat(fmt));
            xls.SetCellValue(23, 6, "Pesos  (Mexico)");
            xls.SetCellValue(24, 2, 1);
            xls.SetCellValue(24, 3, 21.92);
            xls.SetCellValue(24, 4, 3.16);
            xls.SetCellValue(24, 5, 2765);
            xls.SetCellValue(24, 6, 18.21);

            fmt = xls.GetCellVisibleFormatDef(26, 2);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0x33, 0x66, 0xFF);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(26, 2, xls.AddFormat(fmt));
            xls.SetCellValue(26, 2, "Capacidad");

            fmt = xls.GetCellVisibleFormatDef(27, 2);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent4, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(27, 2, xls.AddFormat(fmt));
            xls.SetCellValue(27, 2, "Litro");
            xls.SetCellValue(27, 3, "Galones USA int.");
            xls.SetCellValue(27, 8, "Ojo pues hay dos medidas de Galón (usa-international, y UK)");
            xls.SetCellValue(28, 2, 1);

            fmt = xls.GetCellVisibleFormatDef(28, 3);
            fmt.Format = "0.00";
            xls.SetCellFormat(28, 3, xls.AddFormat(fmt));
            xls.SetCellValue(28, 3, 0.264172051241558);

            fmt = xls.GetCellVisibleFormatDef(30, 2);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent4, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(30, 2, xls.AddFormat(fmt));
            xls.SetCellValue(30, 2, "Galón USA int.");
            xls.SetCellValue(30, 4, "Litro");
            xls.SetCellValue(31, 2, 1);

            fmt = xls.GetCellVisibleFormatDef(31, 4);
            fmt.Format = "0.0";
            xls.SetCellFormat(31, 4, xls.AddFormat(fmt));
            xls.SetCellValue(31, 4, 3.7854118);

            fmt = xls.GetCellVisibleFormatDef(33, 2);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.WrapText = true;
            xls.SetCellFormat(33, 2, xls.AddFormat(fmt));
            xls.SetCellValue(33, 2, "Factores Rendimiento");

            fmt = xls.GetCellVisibleFormatDef(34, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent4, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.WrapText = true;
            xls.SetCellFormat(34, 2, xls.AddFormat(fmt));
            xls.SetCellValue(34, 2, "Cereza a pergamino");
            xls.SetCellValue(34, 3, new TFormula("=$G$59"));

            fmt = xls.GetCellVisibleFormatDef(35, 2);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.WrapText = true;
            xls.SetCellFormat(35, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(36, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent4, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.WrapText = true;
            xls.SetCellFormat(36, 2, xls.AddFormat(fmt));
            xls.SetCellValue(36, 2, "Pergamino a oro");
            xls.SetCellValue(36, 3, new TFormula("=G64"));
            xls.SetCellValue(36, 8, new TFormula("=H44*C11"));

            fmt = xls.GetCellVisibleFormatDef(37, 2);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.WrapText = true;
            xls.SetCellFormat(37, 2, xls.AddFormat(fmt));
            xls.SetCellValue(39, 2, "Conversiones adicionales que salieron de la recolección de informacion");
            xls.SetCellValue(41, 18, "Oro");
            xls.SetCellValue(41, 19, "Pergamino seco");

            fmt = xls.GetCellVisibleFormatDef(42, 2);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(42, 2, xls.AddFormat(fmt));
            xls.SetCellValue(42, 2, "GENERALES");

            fmt = xls.GetCellVisibleFormatDef(42, 3);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(42, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(42, 8);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(42, 8, xls.AddFormat(fmt));
            xls.SetCellValue(42, 8, "Quintal");

            fmt = xls.GetCellVisibleFormatDef(42, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(42, 9, xls.AddFormat(fmt));
            xls.SetCellValue(42, 12, "Un quital de oro entre 17 y 19 latas ");
            xls.SetCellValue(42, 15, "(Gustavo Cerna)");
            xls.SetCellValue(42, 18, new TFormula("=H44"));
            xls.SetCellValue(42, 19, new TFormula("=R42/C36"));

            fmt = xls.GetCellVisibleFormatDef(43, 8);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(43, 8, xls.AddFormat(fmt));
            xls.SetCellValue(43, 8, "Kilos");

            fmt = xls.GetCellVisibleFormatDef(43, 9);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(43, 9, xls.AddFormat(fmt));
            xls.SetCellValue(43, 12, "Una lata entre 7500 - 8500 uvas");
            xls.SetCellValue(44, 2, "1 quintal son 56 kilos");
            xls.SetCellValue(44, 4, "NO son 45.6");
            xls.SetCellValue(44, 5, "Confirmar");

            fmt = xls.GetCellVisibleFormatDef(44, 8);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent4, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(44, 8, xls.AddFormat(fmt));
            xls.SetCellValue(44, 8, 45.6);

            fmt = xls.GetCellVisibleFormatDef(44, 9);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(44, 9, xls.AddFormat(fmt));
            xls.SetCellValue(44, 9, "100 lb");
            xls.SetCellValue(46, 2, "80 quintales de pergamino seco por hectarea en una cosecha");
            xls.SetCellValue(48, 2, "De 20 latas de cereza salen 56 kilos de pergamino seco. Es decir 20 latas de cereza"
            + " por quintal del pergamino seco");
            xls.SetCellValue(49, 2, "Contrastar con P24, P22 que dice 20 latas 15 kilos, parece mas compatible con lo que"
            + " se encuentra en otras referencias");
            xls.SetCellValue(51, 2, "1 lata son 3 kilos (Guia Julio 4)");

            fmt = xls.GetCellVisibleFormatDef(51, 6);
            fmt.WrapText = true;
            xls.SetCellFormat(51, 6, xls.AddFormat(fmt));
            xls.SetCellValue(51, 6, "20 latas cereza");

            fmt = xls.GetCellVisibleFormatDef(51, 7);
            fmt.WrapText = true;
            xls.SetCellFormat(51, 7, xls.AddFormat(fmt));
            xls.SetCellValue(51, 7, "1 quintal pergamino");
            xls.SetCellValue(51, 8, "Factor");

            fmt = xls.GetCellVisibleFormatDef(52, 5);
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(52, 5, xls.AddFormat(fmt));
            xls.SetCellValue(52, 5, "kilos");
            xls.SetCellValue(52, 6, new TFormula("=20*3"));
            xls.SetCellValue(52, 7, 15);

            fmt = xls.GetCellVisibleFormatDef(52, 8);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent4, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(52, 8, xls.AddFormat(fmt));
            xls.SetCellValue(52, 8, new TFormula("=G52/F52"));

            fmt = xls.GetCellVisibleFormatDef(54, 2);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(54, 2, xls.AddFormat(fmt));
            xls.SetCellValue(54, 2, "FACTORES RENDIMIENTO");

            fmt = xls.GetCellVisibleFormatDef(54, 3);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(54, 3, xls.AddFormat(fmt));
            xls.SetCellValue(56, 2, "Cereza a pergamino");
            xls.SetCellValue(56, 4, "Empirico");
            xls.SetCellValue(56, 5, new TFormula("=H52"));
            xls.SetCellValue(56, 10, "(contrastar con P24, P22 que dice 20 latas 15 kilos, parece mas compatible con lo"
            + " que se encuentra en otras referencias)");
            xls.SetCellValue(57, 4, "internet Mexico");
            xls.SetCellValue(57, 5, new TFormula("=E110"));
            xls.SetCellValue(58, 4, "FNC");
            xls.SetCellValue(58, 5, new TFormula("=J158"));
            xls.SetCellValue(59, 4, "COMSA");
            xls.SetCellValue(59, 5, new TFormula("=$E$110"));
            xls.SetCellValue(59, 7, new TFormula("=AVERAGE(E56:E59)"));
            xls.SetCellValue(61, 2, "Pergamino a oro");
            xls.SetCellValue(61, 4, "Prior");
            xls.SetCellValue(61, 5, 0.75);
            xls.SetCellValue(62, 4, "Internet Mexico");
            xls.SetCellValue(62, 5, new TFormula("=$E$111"));
            xls.SetCellValue(63, 4, "FNC");
            xls.SetCellValue(63, 5, new TFormula("=$J$159"));
            xls.SetCellValue(64, 4, "COMSA");
            xls.SetCellValue(64, 5, new TFormula("=$E$105"));
            xls.SetCellValue(64, 7, new TFormula("=AVERAGE(E61:E64)"));

            fmt = xls.GetCellVisibleFormatDef(67, 2);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(67, 2, xls.AddFormat(fmt));
            xls.SetCellValue(67, 2, "VIVERO");

            fmt = xls.GetCellVisibleFormatDef(67, 3);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(67, 3, xls.AddFormat(fmt));
            xls.SetCellValue(69, 2, "Cada peon llena 500 bolsas en un dia");

            fmt = xls.GetCellVisibleFormatDef(72, 2);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(72, 2, xls.AddFormat(fmt));
            xls.SetCellValue(72, 2, "SIEMBRA");

            fmt = xls.GetCellVisibleFormatDef(72, 3);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(72, 3, xls.AddFormat(fmt));
            xls.SetCellValue(74, 2, "10 dias con cerrucho equivalen a 1 con motosierra");
            xls.SetCellValue(74, 6, "El costo de desyerbar una hectarea con motosierra es de 140 soles por hectarea");
            xls.SetCellValue(76, 2, "El desyerbe y la quitada de los hijos para muchos productores es lo mismo");

            fmt = xls.GetCellVisibleFormatDef(79, 2);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(79, 2, xls.AddFormat(fmt));
            xls.SetCellValue(79, 2, "RECOLECCION");

            fmt = xls.GetCellVisibleFormatDef(79, 3);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(79, 3, xls.AddFormat(fmt));
            xls.SetCellValue(80, 4, "Kg");
            xls.SetCellValue(81, 3, "Lata");
            xls.SetCellValue(81, 4, 14);
            xls.SetCellValue(84, 4, "Cereza");
            xls.SetCellValue(84, 5, "Pergamino seco");
            xls.SetCellValue(84, 7, "20 de cereza para una de pergamino seco");
            xls.SetCellValue(84, 11, "(mas bien para un quintal, verificar)");
            xls.SetCellValue(85, 3, "Numero de Latas");
            xls.SetCellValue(85, 4, 20);
            xls.SetCellValue(85, 5, 1);
            xls.SetCellValue(88, 3, "Para la recoleccion");
            xls.SetCellValue(89, 4, "Soles");
            xls.SetCellValue(90, 3, "Pago lata");
            xls.SetCellValue(90, 4, 6);
            xls.SetCellValue(90, 7, "Latas en un día");
            xls.SetCellValue(90, 9, "4 latas");
            xls.SetCellValue(90, 10, "(un peon)");
            xls.SetCellValue(91, 9, "5 latas");
            xls.SetCellValue(96, 2, "Si de una cosecha salen 80 quintales de pergamino seco");
            xls.SetCellValue(96, 7, "[ 80 quintales * (20 latas/quintal) ] /  (5latas/jornada)");
            xls.SetCellValue(98, 7, "quintales");
            xls.SetCellValue(98, 8, "latas");
            xls.SetCellValue(98, 9, "jornada");
            xls.SetCellValue(98, 11, new TFormula("=80*20/5"));
            xls.SetCellValue(98, 12, "320  jornadas para toda la cosecha");
            xls.SetCellValue(99, 8, "quintal");
            xls.SetCellValue(99, 9, "latas");
            xls.SetCellValue(101, 2, "Factores rendimiento COMSA");
            xls.SetCellValue(103, 3, "Uva");
            xls.SetCellValue(103, 4, 615);
            xls.SetCellValue(104, 3, "Pergamino");
            xls.SetCellValue(104, 4, 120);
            xls.SetCellValue(104, 5, new TFormula("=D104/D103"));
            xls.SetCellValue(105, 3, "Oro");
            xls.SetCellValue(105, 4, 100);
            xls.SetCellValue(105, 5, new TFormula("=D105/D104"));
            xls.SetCellValue(105, 7, "Uva a oro");
            xls.SetCellValue(105, 8, 6.15);
            xls.SetCellValue(105, 10, new TFormula("=D103/D105"));
            xls.SetCellValue(107, 2, "Fuentes Internet factores rendimiento");
            xls.SetCellValue(109, 3, "Cereza");
            xls.SetCellValue(109, 4, 250);
            xls.SetCellValue(110, 3, "Pergamino seco");
            xls.SetCellValue(110, 4, 57.5);
            xls.SetCellValue(110, 5, new TFormula("=D110/D109"));
            xls.SetCellValue(111, 3, "Oro");
            xls.SetCellValue(111, 4, 46);
            xls.SetCellValue(111, 5, new TFormula("=D111/D110"));
            xls.SetCellValue(150, 2, "Federacion Nacional");
            xls.SetCellValue(156, 9, "Kilos");
            xls.SetCellValue(157, 8, "Cereza");
            xls.SetCellValue(157, 9, 100);
            xls.SetCellValue(158, 8, "Pergamino");
            xls.SetCellValue(158, 9, new TFormula("=19.5"));
            xls.SetCellValue(158, 10, new TFormula("=I158/I157"));
            xls.SetCellValue(159, 8, "Oro");
            xls.SetCellValue(159, 9, 15.5);
            xls.SetCellValue(159, 10, new TFormula("=I159/I158"));

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
            xls.SetComment(14, 3, new TRichString("Juan Hernandez:\nChristopher Bacon \"Estudio de Costos y Propuesta de Precios para"
            + " sostener el Café, las familias de Productores y Organizaciones Certificadas por Comercio"
            + " Justo en America Latina\"", Runs, xls));

            //You probably don't need to call the lines below. This code is needed only if you want to change the comment box properties like color or default location
            TCommentProperties CommentProps = TCommentProperties.CreateStandard(14, 3, xls);
            CommentProps.Anchor = new TClientAnchor(TFlxAnchorType.DontMoveAndDontResize, 14, 194, 5, 688, 18, 24, 11, 465);

            //Excel by doesn't autofit the comment box so it can hold all text.
            //There is an option in TCommentProperties, but if you use it Excel will show the text in a single line.
            //To have FlexCel autofit the comment for you, you can do it with the following code:

            //    CommentProps.Anchor = xls.AutofitComment(new TRichString("Juan Hernandez:\nChristopher Bacon \"Estudio de Costos y Propuesta de Precios para"
            //+" sostener el Café, las familias de Productores y Organizaciones Certificadas por Comercio"
            //+ " Justo en America Latina\"", Runs, xls), 1.5, true, 1.1, 0, CommentProps.Anchor);

            xls.SetCommentProperties(14, 3, CommentProps);

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
            xls.SetComment(42, 8, new TRichString("Juan Hernandez:\nesto fue confirmado con Gustavo Cerna", Runs, xls));

            //You probably don't need to call the lines below. This code is needed only if you want to change the comment box properties like color or default location
            CommentProps = TCommentProperties.CreateStandard(42, 8, xls);
            CommentProps.Anchor = new TClientAnchor(TFlxAnchorType.DontMoveAndDontResize, 42, 49, 6, 303, 46, 97, 8, 0);

            //Excel by doesn't autofit the comment box so it can hold all text.
            //There is an option in TCommentProperties, but if you use it Excel will show the text in a single line.
            //To have FlexCel autofit the comment for you, you can do it with the following code:

            //    CommentProps.Anchor = xls.AutofitComment(new TRichString("Juan Hernandez:\nesto fue confirmado con Gustavo Cerna", Runs, xls), 1.5, true, 1.1, 0, CommentProps.Anchor);

            xls.SetCommentProperties(42, 8, CommentProps);

            Runs = new TRTFRun[2];
            Runs[0].FirstChar = 0;
            fnt = xls.GetDefaultFont;
            fnt.Size20 = 180;
            fnt.Color = TExcelColor.Automatic;
            fnt.Style = TFlxFontStyles.Bold;
            fnt.Scheme = TFontScheme.None;
            Runs[0].FontIndex = xls.AddFont(fnt);
            Runs[1].FirstChar = 19;
            fnt = xls.GetDefaultFont;
            fnt.Size20 = 180;
            fnt.Color = TExcelColor.Automatic;
            fnt.Scheme = TFontScheme.None;
            Runs[1].FontIndex = xls.AddFont(fnt);
            xls.SetComment(105, 8, new TRichString("Cornell University:\nVer notas cuaderno\nFair Trade", Runs, xls));

            //You probably don't need to call the lines below. This code is needed only if you want to change the comment box properties like color or default location
            CommentProps = TCommentProperties.CreateStandard(105, 8, xls);
            CommentProps.Anchor = new TClientAnchor(TFlxAnchorType.DontMoveAndDontResize, 104, 49, 9, 163, 108, 97, 10, 628);

            //Excel by doesn't autofit the comment box so it can hold all text.
            //There is an option in TCommentProperties, but if you use it Excel will show the text in a single line.
            //To have FlexCel autofit the comment for you, you can do it with the following code:

            //    CommentProps.Anchor = xls.AutofitComment(new TRichString("Cornell University:\nVer notas cuaderno\nFair Trade", Runs, xls), 1.5, true, 1.1, 0, CommentProps.Anchor);

            xls.SetCommentProperties(105, 8, CommentProps);

            //Cell selection and scroll position.
            xls.SelectCell(24, 6, false);
            xls.ScrollWindow(2, 1);

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
