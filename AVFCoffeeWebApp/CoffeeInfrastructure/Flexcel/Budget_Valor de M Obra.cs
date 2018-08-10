using CoffeeCore.Interfaces;
using FlexCel.Core;
using FlexCel.XlsAdapter;
using System;
using System.Collections.Generic;
using System.IO;

namespace CoffeeInfrastructure.Flexcel
{
    public class Budget_Valor_de_M_Obra
    {

        public void Budget_Valor_M_De_Obra(ExcelFile xls)
        {
            xls.NewFile(34, TExcelFileFormat.v2016);    //Create a new Excel file with 34 sheets.

            //Set the names of the sheets
            xls.ActiveSheet = 1;
            xls.SheetName = "Metrics";
            xls.ActiveSheet = 2;
            xls.SheetName = "Inputs 1.0";
            xls.ActiveSheet = 3;
            xls.SheetName = "Inputs advance 2.0 (eng)";
            xls.ActiveSheet = 4;
            xls.SheetName = "Outcome 1.0";
            xls.ActiveSheet = 5;
            xls.SheetName = "Additional 2.0";
            xls.ActiveSheet = 6;
            xls.SheetName = "Fixed 2.0";
            xls.ActiveSheet = 7;
            xls.SheetName = "Variable 2.0";
            xls.ActiveSheet = 8;
            xls.SheetName = "General Budget 2.0";
            xls.ActiveSheet = 9;
            xls.SheetName = "DATABASE_Schema";
            xls.ActiveSheet = 10;
            xls.SheetName = "Inputs 2.0 Conv. default values";
            xls.ActiveSheet = 11;
            xls.SheetName = "Inputs 2.0 Conv. new inputs";
            xls.ActiveSheet = 12;
            xls.SheetName = "Inputs advanced 2.0 (esp_eng)";
            xls.ActiveSheet = 13;
            xls.SheetName = "Inputs TOT advanced";
            xls.ActiveSheet = 14;
            xls.SheetName = "Gral Conf. Summary";
            xls.ActiveSheet = 15;
            xls.SheetName = "Inputs 1.0 default values";
            xls.ActiveSheet = 16;
            xls.SheetName = "Inputs 1.0 Conv. new values";
            xls.ActiveSheet = 17;
            xls.SheetName = "Outcome TOTAL_Adj";
            xls.ActiveSheet = 18;
            xls.SheetName = "Outcome_Y_Adjustment";
            xls.ActiveSheet = 19;
            xls.SheetName = "Outcome_L Adjustment";
            xls.ActiveSheet = 20;
            xls.SheetName = "Proportions";
            xls.ActiveSheet = 21;
            xls.SheetName = "Budget_Supuestos";
            xls.ActiveSheet = 22;
            xls.SheetName = "Budget_Equipo";
            xls.ActiveSheet = 23;
            xls.SheetName = "Budget_M Obra";
            xls.ActiveSheet = 24;
            xls.SheetName = "Budget_M Obra_1";
            xls.ActiveSheet = 25;
            xls.SheetName = "Budget_Valor de M Obra_1";
            xls.ActiveSheet = 26;
            xls.SheetName = "Budget_Establecimiento_1";
            xls.ActiveSheet = 27;
            xls.SheetName = "Budget_Presupuesto";
            xls.ActiveSheet = 28;
            xls.SheetName = "Budget_Valor de M Obra";
            xls.ActiveSheet = 29;
            xls.SheetName = "Budget_Establecimiento";
            xls.ActiveSheet = 30;
            xls.SheetName = "Budget_Sostenemiento";
            xls.ActiveSheet = 31;
            xls.SheetName = "Outcome 1.0 pre_metric_currency";
            xls.ActiveSheet = 32;
            xls.SheetName = "Conversiones";
            xls.ActiveSheet = 33;
            xls.SheetName = "Proporción de productividad";
            xls.ActiveSheet = 34;
            xls.SheetName = "Inputs 1.0 (Ref)";

            xls.ActiveSheet = 25;    //Set the sheet we are working in.

            //Global Workbook Options
            xls.OptionsAutoCompressPictures = false;
            xls.OptionsMultithreadRecalc = 0;

            //Sheet Options
            xls.SheetName = "Budget_Valor de M Obra_1";

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
            Range = new TXlsNamedRange(RangeName, 29, 32, "=Budget_Establecimiento!$A$3:$C$53");
            //You could also use: Range = new TXlsNamedRange(RangeName, 29, 29, 3, 1, 53, 3, 32);
            xls.SetNamedRange(Range);

            RangeName = TXlsNamedRange.GetInternalName(InternalNameRange.Print_Area);
            Range = new TXlsNamedRange(RangeName, 23, 32, "='Budget_M Obra'!$A$1:$K$86");
            //You could also use: Range = new TXlsNamedRange(RangeName, 23, 23, 1, 1, 86, 11, 32);
            xls.SetNamedRange(Range);

            RangeName = TXlsNamedRange.GetInternalName(InternalNameRange.Print_Area);
            Range = new TXlsNamedRange(RangeName, 27, 32, "=Budget_Presupuesto!$A$34:$J$46");
            //You could also use: Range = new TXlsNamedRange(RangeName, 27, 27, 34, 1, 46, 10, 32);
            xls.SetNamedRange(Range);

            RangeName = TXlsNamedRange.GetInternalName(InternalNameRange.Print_Area);
            Range = new TXlsNamedRange(RangeName, 30, 32, "=Budget_Sostenemiento!$A$1:$K$44");
            //You could also use: Range = new TXlsNamedRange(RangeName, 30, 30, 1, 1, 44, 11, 32);
            xls.SetNamedRange(Range);

            RangeName = TXlsNamedRange.GetInternalName(InternalNameRange.Print_Area);
            Range = new TXlsNamedRange(RangeName, 21, 32, "=Budget_Supuestos!$A$276:$G$297");
            //You could also use: Range = new TXlsNamedRange(RangeName, 21, 21, 276, 1, 297, 7, 32);
            xls.SetNamedRange(Range);

            RangeName = TXlsNamedRange.GetInternalName(InternalNameRange.Print_Area);
            Range = new TXlsNamedRange(RangeName, 28, 32, "='Budget_Valor de M Obra'!$A$2:$J$85");
            //You could also use: Range = new TXlsNamedRange(RangeName, 28, 28, 2, 1, 85, 10, 32);
            xls.SetNamedRange(Range);


            //Printer Settings
            xls.PrintOptions = TPrintOptions.Orientation | TPrintOptions.NoPls;

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
            xls.DefaultColWidth = 2304;
            xls.DefaultRowHeight = 315;

            xls.SetRowHeight(36, 375);    //18.75 * 20
            xls.SetRowHeight(45, 375);    //18.75 * 20
            xls.SetRowHeight(61, 915);    //45.75 * 20
            xls.SetRowHeight(62, 3315);    //165.75 * 20
            xls.SetRowHeight(85, 5040);    //252.00 * 20
            xls.SetRowHeight(86, 5985);    //299.25 * 20
            xls.SetRowHeight(87, 4725);    //236.25 * 20
            xls.SetRowHeight(88, 6300);    //315.00 * 20

            //Set the cell values
            TFlxFormat fmt;
            fmt = xls.GetCellVisibleFormatDef(2, 1);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(2, 1, xls.AddFormat(fmt));
            xls.SetCellValue(2, 1, "Valor de Mano de Obra");

            fmt = xls.GetCellVisibleFormatDef(3, 1);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(3, 1, xls.AddFormat(fmt));

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
            xls.SetCellValue(5, 1, "Valor Mano de obra para el germinador");

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

            fmt = xls.GetCellVisibleFormatDef(6, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(6, 1, xls.AddFormat(fmt));
            xls.SetCellValue(6, 1, "Recolección de semillas");

            fmt = xls.GetCellVisibleFormatDef(6, 2);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(6, 2, xls.AddFormat(fmt));
            xls.SetCellValue(6, 2, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!B6*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(6, 3);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "0.00";
            xls.SetCellFormat(6, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(6, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(6, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(6, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(6, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(6, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(6, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(6, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(6, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(6, 8);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(6, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(6, 9);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(6, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(6, 10);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
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

            fmt = xls.GetCellVisibleFormatDef(6, 20);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(6, 20, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(6, 21);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(6, 21, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(6, 22);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(6, 22, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(6, 23);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(6, 23, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(6, 24);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(6, 24, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(7, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(7, 1, xls.AddFormat(fmt));
            xls.SetCellValue(7, 1, "Selección de semillas");

            fmt = xls.GetCellVisibleFormatDef(7, 2);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(7, 2, xls.AddFormat(fmt));
            xls.SetCellValue(7, 2, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!B7*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(7, 3);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "0.00";
            xls.SetCellFormat(7, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(7, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(7, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(7, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(7, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(7, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(7, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(7, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(7, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(7, 8);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(7, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(7, 9);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(7, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(7, 10);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
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

            fmt = xls.GetCellVisibleFormatDef(7, 20);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(7, 20, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(7, 21);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(7, 21, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(7, 22);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(7, 22, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(7, 23);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(7, 23, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(7, 24);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(7, 24, xls.AddFormat(fmt));

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
            xls.SetCellFormat(8, 2, xls.AddFormat(fmt));
            xls.SetCellValue(8, 2, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!B8*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(8, 3);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "0.00";
            xls.SetCellFormat(8, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(8, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(8, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(8, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(8, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(8, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(8, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(8, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(8, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(8, 8);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(8, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(8, 9);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(8, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(8, 10);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
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

            fmt = xls.GetCellVisibleFormatDef(8, 20);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(8, 20, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(8, 21);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(8, 21, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(8, 22);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(8, 22, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(8, 23);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(8, 23, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(8, 24);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(8, 24, xls.AddFormat(fmt));

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
            xls.SetCellFormat(9, 2, xls.AddFormat(fmt));
            xls.SetCellValue(9, 2, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!B9*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(9, 3);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "0.00";
            xls.SetCellFormat(9, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(9, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(9, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(9, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(9, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(9, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(9, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(9, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(9, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(9, 8);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(9, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(9, 9);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(9, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(9, 10);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
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

            fmt = xls.GetCellVisibleFormatDef(9, 20);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(9, 20, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(9, 21);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(9, 21, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(9, 22);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(9, 22, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(9, 23);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(9, 23, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(9, 24);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(9, 24, xls.AddFormat(fmt));

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
            xls.SetCellFormat(10, 2, xls.AddFormat(fmt));
            xls.SetCellValue(10, 2, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!B10*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(10, 3);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "0.00";
            xls.SetCellFormat(10, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(10, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(10, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(10, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(10, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(10, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(10, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(10, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(10, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(10, 8);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(10, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(10, 9);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(10, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(10, 10);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
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

            fmt = xls.GetCellVisibleFormatDef(10, 20);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(10, 20, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(10, 21);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(10, 21, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(10, 22);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(10, 22, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(10, 23);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(10, 23, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(10, 24);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(10, 24, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(11, 1);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Font.Style = TFlxFontStyles.Italic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0x33, 0x66, 0xFF);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(11, 1, xls.AddFormat(fmt));
            xls.SetCellValue(11, 1, "Total Valor Mano Obra Germinador");

            fmt = xls.GetCellVisibleFormatDef(11, 2);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Font.Style = TFlxFontStyles.Bold | TFlxFontStyles.Italic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0x33, 0x66, 0xFF);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(11, 2, xls.AddFormat(fmt));
            xls.SetCellValue(11, 2, new TFormula("=SUM(B6:B10)"));

            fmt = xls.GetCellVisibleFormatDef(11, 3);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "0.00";
            xls.SetCellFormat(11, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(11, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(11, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(11, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(11, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(11, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(11, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(11, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(11, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(11, 8);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(11, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(11, 9);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(11, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(11, 10);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
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

            fmt = xls.GetCellVisibleFormatDef(11, 20);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(11, 20, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(11, 21);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(11, 21, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(11, 22);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(11, 22, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(11, 23);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(11, 23, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(11, 24);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(11, 24, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(12, 1);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(12, 1, xls.AddFormat(fmt));
            xls.SetCellValue(12, 1, "Valor Mano de obra para el vivero");

            fmt = xls.GetCellVisibleFormatDef(12, 2);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(12, 2, xls.AddFormat(fmt));
            xls.SetCellValue(12, 2, "Año 0");

            fmt = xls.GetCellVisibleFormatDef(12, 3);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(12, 3, xls.AddFormat(fmt));
            xls.SetCellValue(12, 3, "Año 1");

            fmt = xls.GetCellVisibleFormatDef(12, 4);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(12, 4, xls.AddFormat(fmt));
            xls.SetCellValue(12, 4, "Año 2");

            fmt = xls.GetCellVisibleFormatDef(12, 5);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(12, 5, xls.AddFormat(fmt));
            xls.SetCellValue(12, 5, "Año 3");

            fmt = xls.GetCellVisibleFormatDef(12, 6);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(12, 6, xls.AddFormat(fmt));
            xls.SetCellValue(12, 6, "Año 4");

            fmt = xls.GetCellVisibleFormatDef(12, 7);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(12, 7, xls.AddFormat(fmt));
            xls.SetCellValue(12, 7, "Año 5");

            fmt = xls.GetCellVisibleFormatDef(12, 8);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(12, 8, xls.AddFormat(fmt));
            xls.SetCellValue(12, 8, "Año 6");

            fmt = xls.GetCellVisibleFormatDef(12, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(12, 9, xls.AddFormat(fmt));
            xls.SetCellValue(12, 9, "Año 7");

            fmt = xls.GetCellVisibleFormatDef(12, 10);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
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

            fmt = xls.GetCellVisibleFormatDef(12, 20);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(12, 20, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(12, 21);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(12, 21, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(12, 22);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(12, 22, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(12, 23);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(12, 23, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(12, 24);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(12, 24, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(13, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(13, 1, xls.AddFormat(fmt));
            xls.SetCellValue(13, 1, "Construcción del vivero");

            fmt = xls.GetCellVisibleFormatDef(13, 2);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(13, 2, xls.AddFormat(fmt));
            xls.SetCellValue(13, 2, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!B13*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(13, 3);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(13, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(13, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(13, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(13, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(13, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(13, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(13, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(13, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(13, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(13, 8);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(13, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(13, 9);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(13, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(13, 10);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
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

            fmt = xls.GetCellVisibleFormatDef(13, 20);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(13, 20, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(13, 21);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(13, 21, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(13, 22);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(13, 22, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(13, 23);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(13, 23, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(13, 24);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(13, 24, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(14, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(14, 1, xls.AddFormat(fmt));
            xls.SetCellValue(14, 1, "Jalada y arrancada de la tierra para el vivero");

            fmt = xls.GetCellVisibleFormatDef(14, 2);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(14, 2, xls.AddFormat(fmt));
            xls.SetCellValue(14, 2, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!B14*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(14, 3);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(14, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(14, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(14, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(14, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(14, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(14, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(14, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(14, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(14, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(14, 8);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(14, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(14, 9);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(14, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(14, 10);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
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

            fmt = xls.GetCellVisibleFormatDef(14, 20);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(14, 20, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(14, 21);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(14, 21, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(14, 22);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(14, 22, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(14, 23);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(14, 23, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(14, 24);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(14, 24, xls.AddFormat(fmt));

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
            xls.SetCellFormat(15, 2, xls.AddFormat(fmt));
            xls.SetCellValue(15, 2, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!B15*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(15, 3);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(15, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(15, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(15, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(15, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(15, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(15, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(15, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(15, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(15, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(15, 8);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(15, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(15, 9);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(15, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(15, 10);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
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

            fmt = xls.GetCellVisibleFormatDef(15, 20);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(15, 20, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(15, 21);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(15, 21, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(15, 22);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(15, 22, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(15, 23);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(15, 23, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(15, 24);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(15, 24, xls.AddFormat(fmt));

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
            xls.SetCellFormat(16, 2, xls.AddFormat(fmt));
            xls.SetCellValue(16, 2, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!B16*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(16, 3);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(16, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(16, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(16, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(16, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(16, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(16, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(16, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(16, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(16, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(16, 8);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(16, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(16, 9);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(16, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(16, 10);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
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

            fmt = xls.GetCellVisibleFormatDef(16, 20);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(16, 20, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(16, 21);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(16, 21, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(16, 22);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(16, 22, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(16, 23);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(16, 23, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(16, 24);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(16, 24, xls.AddFormat(fmt));

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
            xls.SetCellFormat(17, 2, xls.AddFormat(fmt));
            xls.SetCellValue(17, 2, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!B17*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(17, 3);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(17, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(17, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(17, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(17, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(17, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(17, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(17, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(17, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(17, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(17, 8);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(17, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(17, 9);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(17, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(17, 10);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
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

            fmt = xls.GetCellVisibleFormatDef(17, 20);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(17, 20, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(17, 21);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(17, 21, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(17, 22);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(17, 22, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(17, 23);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(17, 23, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(17, 24);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(17, 24, xls.AddFormat(fmt));

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
            xls.SetCellFormat(18, 2, xls.AddFormat(fmt));
            xls.SetCellValue(18, 2, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!B18*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(18, 3);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(18, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(18, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(18, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(18, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(18, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(18, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(18, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(18, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(18, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(18, 8);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(18, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(18, 9);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(18, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(18, 10);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
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

            fmt = xls.GetCellVisibleFormatDef(18, 20);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(18, 20, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(18, 21);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(18, 21, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(18, 22);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(18, 22, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(18, 23);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(18, 23, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(18, 24);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(18, 24, xls.AddFormat(fmt));

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
            xls.SetCellFormat(19, 2, xls.AddFormat(fmt));
            xls.SetCellValue(19, 2, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!B19*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(19, 3);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(19, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(19, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(19, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(19, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(19, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(19, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(19, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(19, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(19, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(19, 8);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(19, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(19, 9);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(19, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(19, 10);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
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

            fmt = xls.GetCellVisibleFormatDef(19, 20);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(19, 20, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(19, 21);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(19, 21, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(19, 22);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(19, 22, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(19, 23);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(19, 23, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(19, 24);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(19, 24, xls.AddFormat(fmt));

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
            xls.SetCellFormat(20, 2, xls.AddFormat(fmt));
            xls.SetCellValue(20, 2, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!B20*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(20, 3);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(20, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(20, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(20, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(20, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(20, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(20, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(20, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(20, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(20, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(20, 8);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(20, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(20, 9);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(20, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(20, 10);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
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

            fmt = xls.GetCellVisibleFormatDef(20, 20);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(20, 20, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(20, 21);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(20, 21, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(20, 22);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(20, 22, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(20, 23);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(20, 23, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(20, 24);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(20, 24, xls.AddFormat(fmt));

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
            xls.SetCellFormat(21, 2, xls.AddFormat(fmt));
            xls.SetCellValue(21, 2, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!B21*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(21, 3);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(21, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(21, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(21, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(21, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(21, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(21, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(21, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(21, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(21, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(21, 8);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(21, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(21, 9);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(21, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(21, 10);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
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

            fmt = xls.GetCellVisibleFormatDef(21, 20);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(21, 20, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(21, 21);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(21, 21, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(21, 22);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(21, 22, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(21, 23);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(21, 23, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(21, 24);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(21, 24, xls.AddFormat(fmt));

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
            xls.SetCellFormat(22, 2, xls.AddFormat(fmt));
            xls.SetCellValue(22, 2, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!B22*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(22, 3);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent6, 0.399975585192419);
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(22, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(22, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent6, 0.399975585192419);
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(22, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(22, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent6, 0.399975585192419);
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(22, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(22, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent6, 0.399975585192419);
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(22, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(22, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent6, 0.399975585192419);
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(22, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(22, 8);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent6, 0.399975585192419);
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(22, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(22, 9);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent6, 0.399975585192419);
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(22, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(22, 10);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent6, 0.399975585192419);
            fmt.Font.Scheme = TFontScheme.None;
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

            fmt = xls.GetCellVisibleFormatDef(22, 20);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent6, 0.399975585192419);
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(22, 20, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(22, 21);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent6, 0.399975585192419);
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(22, 21, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(22, 22);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent6, 0.399975585192419);
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(22, 22, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(22, 23);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent6, 0.399975585192419);
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(22, 23, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(22, 24);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent6, 0.399975585192419);
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(22, 24, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(23, 1);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Font.Style = TFlxFontStyles.Italic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0x33, 0x66, 0xFF);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(23, 1, xls.AddFormat(fmt));
            xls.SetCellValue(23, 1, "Total Valor Mano Obra vivero");

            fmt = xls.GetCellVisibleFormatDef(23, 2);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Font.Style = TFlxFontStyles.Bold | TFlxFontStyles.Italic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0x33, 0x66, 0xFF);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(23, 2, xls.AddFormat(fmt));
            xls.SetCellValue(23, 2, new TFormula("=SUM(B13:B22)"));

            fmt = xls.GetCellVisibleFormatDef(23, 3);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(23, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(23, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(23, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(23, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(23, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(23, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(23, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(23, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(23, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(23, 8);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(23, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(23, 9);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(23, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(23, 10);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
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

            fmt = xls.GetCellVisibleFormatDef(23, 20);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(23, 20, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(23, 21);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(23, 21, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(23, 22);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(23, 22, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(23, 23);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(23, 23, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(23, 24);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(23, 24, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(24, 1);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(24, 1, xls.AddFormat(fmt));
            xls.SetCellValue(24, 1, "Valor mano de obra preparación terreno para renovacion");

            fmt = xls.GetCellVisibleFormatDef(24, 2);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(24, 2, xls.AddFormat(fmt));
            xls.SetCellValue(24, 2, "Año 0");

            fmt = xls.GetCellVisibleFormatDef(24, 3);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(24, 3, xls.AddFormat(fmt));
            xls.SetCellValue(24, 3, "Año 1");

            fmt = xls.GetCellVisibleFormatDef(24, 4);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(24, 4, xls.AddFormat(fmt));
            xls.SetCellValue(24, 4, "Año 2");

            fmt = xls.GetCellVisibleFormatDef(24, 5);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(24, 5, xls.AddFormat(fmt));
            xls.SetCellValue(24, 5, "Año 3");

            fmt = xls.GetCellVisibleFormatDef(24, 6);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(24, 6, xls.AddFormat(fmt));
            xls.SetCellValue(24, 6, "Año 4");

            fmt = xls.GetCellVisibleFormatDef(24, 7);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(24, 7, xls.AddFormat(fmt));
            xls.SetCellValue(24, 7, "Año 5");

            fmt = xls.GetCellVisibleFormatDef(24, 8);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(24, 8, xls.AddFormat(fmt));
            xls.SetCellValue(24, 8, "Año 6");

            fmt = xls.GetCellVisibleFormatDef(24, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(24, 9, xls.AddFormat(fmt));
            xls.SetCellValue(24, 9, "Año 7");

            fmt = xls.GetCellVisibleFormatDef(24, 10);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(24, 10, xls.AddFormat(fmt));
            xls.SetCellValue(24, 10, "Año 8");

            fmt = xls.GetCellVisibleFormatDef(25, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(25, 1, xls.AddFormat(fmt));
            xls.SetCellValue(25, 1, "Limpia del terreno");

            fmt = xls.GetCellVisibleFormatDef(25, 2);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(25, 2, xls.AddFormat(fmt));
            xls.SetCellValue(25, 2, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!B25*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(25, 3);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(25, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(25, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(25, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(25, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(25, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(25, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(25, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(25, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(25, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(25, 8);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(25, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(25, 9);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(25, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(25, 10);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
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

            fmt = xls.GetCellVisibleFormatDef(25, 20);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(25, 20, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(25, 21);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(25, 21, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(25, 22);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(25, 22, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(25, 23);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(25, 23, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(25, 24);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(25, 24, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(26, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(26, 1, xls.AddFormat(fmt));
            xls.SetCellValue(26, 1, "Corte de arboles de café viejos u otros maderables");

            fmt = xls.GetCellVisibleFormatDef(26, 2);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Underline = TFlxUnderline.Single;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(26, 2, xls.AddFormat(fmt));
            xls.SetCellValue(26, 2, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!B26*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(26, 3);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(26, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(26, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(26, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(26, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(26, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(26, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(26, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(26, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(26, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(26, 8);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(26, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(26, 9);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(26, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(26, 10);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
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

            fmt = xls.GetCellVisibleFormatDef(26, 20);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(26, 20, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(26, 21);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(26, 21, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(26, 22);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(26, 22, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(26, 23);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(26, 23, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(26, 24);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(26, 24, xls.AddFormat(fmt));

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
            xls.SetCellFormat(27, 2, xls.AddFormat(fmt));
            xls.SetCellValue(27, 2, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!B27*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(27, 3);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(27, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(27, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(27, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(27, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(27, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(27, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(27, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(27, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(27, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(27, 8);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(27, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(27, 9);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(27, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(27, 10);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
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

            fmt = xls.GetCellVisibleFormatDef(27, 20);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(27, 20, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(27, 21);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(27, 21, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(27, 22);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(27, 22, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(27, 23);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(27, 23, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(27, 24);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(27, 24, xls.AddFormat(fmt));

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
            xls.SetCellFormat(28, 2, xls.AddFormat(fmt));
            xls.SetCellValue(28, 2, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!B28*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(28, 3);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(28, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(28, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(28, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(28, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(28, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(28, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(28, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(28, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(28, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(28, 8);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(28, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(28, 9);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(28, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(28, 10);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
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

            fmt = xls.GetCellVisibleFormatDef(28, 20);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(28, 20, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(28, 21);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(28, 21, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(28, 22);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(28, 22, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(28, 23);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(28, 23, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(28, 24);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(28, 24, xls.AddFormat(fmt));

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
            xls.SetCellFormat(29, 2, xls.AddFormat(fmt));
            xls.SetCellValue(29, 2, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!B29*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(29, 3);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(29, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(29, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(29, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(29, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(29, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(29, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(29, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(29, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(29, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(29, 8);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(29, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(29, 9);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(29, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(29, 10);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
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

            fmt = xls.GetCellVisibleFormatDef(29, 20);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(29, 20, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(29, 21);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(29, 21, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(29, 22);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(29, 22, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(29, 23);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(29, 23, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(29, 24);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(29, 24, xls.AddFormat(fmt));

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
            xls.SetCellFormat(30, 2, xls.AddFormat(fmt));
            xls.SetCellValue(30, 2, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!B30*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(30, 3);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(30, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(30, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(30, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(30, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(30, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(30, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(30, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(30, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(30, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(30, 8);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(30, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(30, 9);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(30, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(30, 10);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
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

            fmt = xls.GetCellVisibleFormatDef(30, 20);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(30, 20, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(30, 21);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(30, 21, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(30, 22);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(30, 22, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(30, 23);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(30, 23, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(30, 24);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(30, 24, xls.AddFormat(fmt));

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
            xls.SetCellFormat(31, 2, xls.AddFormat(fmt));
            xls.SetCellValue(31, 2, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!B31*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(31, 3);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(31, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(31, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(31, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(31, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(31, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(31, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(31, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(31, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(31, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(31, 8);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(31, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(31, 9);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(31, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(31, 10);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
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

            fmt = xls.GetCellVisibleFormatDef(31, 20);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(31, 20, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(31, 21);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(31, 21, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(31, 22);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(31, 22, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(31, 23);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(31, 23, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(31, 24);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(31, 24, xls.AddFormat(fmt));

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
            xls.SetCellFormat(32, 2, xls.AddFormat(fmt));
            xls.SetCellValue(32, 2, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!B32*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(32, 3);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(32, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(32, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(32, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(32, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(32, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(32, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(32, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(32, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(32, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(32, 8);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(32, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(32, 9);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(32, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(32, 10);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
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

            fmt = xls.GetCellVisibleFormatDef(32, 20);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(32, 20, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(32, 21);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(32, 21, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(32, 22);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(32, 22, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(32, 23);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(32, 23, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(32, 24);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(32, 24, xls.AddFormat(fmt));

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
            xls.SetCellFormat(33, 2, xls.AddFormat(fmt));
            xls.SetCellValue(33, 2, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!B33*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(33, 3);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(33, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(33, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(33, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(33, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(33, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(33, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(33, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(33, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(33, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(33, 8);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(33, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(33, 9);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(33, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(33, 10);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(33, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(33, 11);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(33, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(33, 12);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(33, 12, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(33, 13);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(33, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(33, 14);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(33, 14, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(33, 15);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(33, 15, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(33, 16);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(33, 16, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(33, 17);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(33, 17, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(33, 18);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(33, 18, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(33, 19);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(33, 19, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(33, 20);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(33, 20, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(33, 21);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(33, 21, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(33, 22);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(33, 22, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(33, 23);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(33, 23, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(33, 24);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(33, 24, xls.AddFormat(fmt));

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
            xls.SetCellFormat(34, 2, xls.AddFormat(fmt));
            xls.SetCellValue(34, 2, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!B34*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(34, 3);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(34, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(34, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(34, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(34, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(34, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(34, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(34, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(34, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(34, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(34, 8);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(34, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(34, 9);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(34, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(34, 10);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(34, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(34, 11);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(34, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(34, 12);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(34, 12, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(34, 13);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(34, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(34, 14);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(34, 14, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(34, 15);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(34, 15, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(34, 16);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(34, 16, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(34, 17);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(34, 17, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(34, 18);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(34, 18, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(34, 19);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(34, 19, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(34, 20);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(34, 20, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(34, 21);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(34, 21, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(34, 22);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(34, 22, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(34, 23);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(34, 23, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(34, 24);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(34, 24, xls.AddFormat(fmt));

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
            xls.SetCellFormat(35, 2, xls.AddFormat(fmt));
            xls.SetCellValue(35, 2, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!B35*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(35, 3);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(35, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(35, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(35, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(35, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(35, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(35, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(35, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(35, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(35, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(35, 8);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(35, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(35, 9);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(35, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(35, 10);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
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

            fmt = xls.GetCellVisibleFormatDef(35, 20);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(35, 20, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(35, 21);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(35, 21, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(35, 22);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(35, 22, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(35, 23);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(35, 23, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(35, 24);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(35, 24, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(36, 1);
            fmt.Font.Size20 = 280;
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Font.Style = TFlxFontStyles.Italic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0x33, 0x66, 0xFF);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(36, 1, xls.AddFormat(fmt));
            xls.SetCellValue(36, 1, "Valor Mano Obra Terreno para Renovación");

            fmt = xls.GetCellVisibleFormatDef(36, 2);
            fmt.Font.Size20 = 280;
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Font.Style = TFlxFontStyles.Bold | TFlxFontStyles.Italic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0x33, 0x66, 0xFF);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(36, 2, xls.AddFormat(fmt));
            xls.SetCellValue(36, 2, new TFormula("=SUM(B25:B35)"));

            fmt = xls.GetCellVisibleFormatDef(37, 1);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Font.Style = TFlxFontStyles.Italic;
            xls.SetCellFormat(37, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(37, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(37, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(38, 1);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(38, 1, xls.AddFormat(fmt));
            xls.SetCellValue(38, 1, "Año 1");

            fmt = xls.GetCellVisibleFormatDef(38, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(38, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(39, 1);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(39, 1, xls.AddFormat(fmt));
            xls.SetCellValue(39, 1, "Valor mano de obra para la plantilla o levante ");

            fmt = xls.GetCellVisibleFormatDef(39, 2);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(39, 2, xls.AddFormat(fmt));
            xls.SetCellValue(39, 2, "Año 0");

            fmt = xls.GetCellVisibleFormatDef(39, 3);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(39, 3, xls.AddFormat(fmt));
            xls.SetCellValue(39, 3, "Año 1");

            fmt = xls.GetCellVisibleFormatDef(39, 4);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(39, 4, xls.AddFormat(fmt));
            xls.SetCellValue(39, 4, "Año 2");

            fmt = xls.GetCellVisibleFormatDef(39, 5);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(39, 5, xls.AddFormat(fmt));
            xls.SetCellValue(39, 5, "Año 3");

            fmt = xls.GetCellVisibleFormatDef(39, 6);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(39, 6, xls.AddFormat(fmt));
            xls.SetCellValue(39, 6, "Año 4");

            fmt = xls.GetCellVisibleFormatDef(39, 7);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(39, 7, xls.AddFormat(fmt));
            xls.SetCellValue(39, 7, "Año 5");

            fmt = xls.GetCellVisibleFormatDef(39, 8);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(39, 8, xls.AddFormat(fmt));
            xls.SetCellValue(39, 8, "Año 6");

            fmt = xls.GetCellVisibleFormatDef(39, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(39, 9, xls.AddFormat(fmt));
            xls.SetCellValue(39, 9, "Año 7");

            fmt = xls.GetCellVisibleFormatDef(39, 10);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(39, 10, xls.AddFormat(fmt));
            xls.SetCellValue(39, 10, "Año 8");

            fmt = xls.GetCellVisibleFormatDef(40, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(40, 1, xls.AddFormat(fmt));
            xls.SetCellValue(40, 1, "Desyerbe periodico ");

            fmt = xls.GetCellVisibleFormatDef(40, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(40, 3, xls.AddFormat(fmt));
            xls.SetCellValue(40, 3, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!C41*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(40, 4);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            xls.SetCellFormat(40, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(41, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(41, 1, xls.AddFormat(fmt));
            xls.SetCellValue(41, 1, "Aplicación de abonos orgánicos para levante");

            fmt = xls.GetCellVisibleFormatDef(41, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(41, 3, xls.AddFormat(fmt));
            xls.SetCellValue(41, 3, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!C42*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(41, 4);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            xls.SetCellFormat(41, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(41, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(41, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(42, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(42, 1, xls.AddFormat(fmt));
            xls.SetCellValue(42, 1, "Aplicación de abonos químicos para levante");

            fmt = xls.GetCellVisibleFormatDef(42, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(42, 3, xls.AddFormat(fmt));
            xls.SetCellValue(42, 3, 0);

            fmt = xls.GetCellVisibleFormatDef(42, 4);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            xls.SetCellFormat(42, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(43, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(43, 1, xls.AddFormat(fmt));
            xls.SetCellValue(43, 1, "Aplicación de foliares para fertilización y control roya");

            fmt = xls.GetCellVisibleFormatDef(43, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(43, 3, xls.AddFormat(fmt));
            xls.SetCellValue(43, 3, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!C44*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(43, 4);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            xls.SetCellFormat(43, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(44, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(44, 1, xls.AddFormat(fmt));
            xls.SetCellValue(44, 1, "Otros:");

            fmt = xls.GetCellVisibleFormatDef(44, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(44, 3, xls.AddFormat(fmt));
            xls.SetCellValue(44, 3, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!C45*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(44, 4);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            xls.SetCellFormat(44, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(45, 1);
            fmt.Font.Size20 = 280;
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Font.Style = TFlxFontStyles.Italic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0x33, 0x66, 0xFF);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(45, 1, xls.AddFormat(fmt));
            xls.SetCellValue(45, 1, "Valor mano obra para la plantilla o el levante");

            fmt = xls.GetCellVisibleFormatDef(45, 2);
            fmt.Font.Size20 = 280;
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Font.Style = TFlxFontStyles.Bold | TFlxFontStyles.Italic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0x33, 0x66, 0xFF);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(45, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(45, 3);
            fmt.Font.Size20 = 280;
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Font.Style = TFlxFontStyles.Italic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0x33, 0x66, 0xFF);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(45, 3, xls.AddFormat(fmt));
            xls.SetCellValue(45, 3, new TFormula("=SUM(C40:C44)"));

            fmt = xls.GetCellVisibleFormatDef(45, 4);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            xls.SetCellFormat(45, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(46, 1);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Font.Style = TFlxFontStyles.Italic;
            xls.SetCellFormat(46, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(46, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(46, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(47, 1);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(47, 1, xls.AddFormat(fmt));
            xls.SetCellValue(47, 1, "Año 2-8");

            fmt = xls.GetCellVisibleFormatDef(47, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(47, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(48, 1);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(48, 1, xls.AddFormat(fmt));
            xls.SetCellValue(48, 1, "Valor mano de obra para mantenimiento");

            fmt = xls.GetCellVisibleFormatDef(48, 2);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(48, 2, xls.AddFormat(fmt));
            xls.SetCellValue(48, 2, "Año 0");

            fmt = xls.GetCellVisibleFormatDef(48, 3);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(48, 3, xls.AddFormat(fmt));
            xls.SetCellValue(48, 3, "Año 1");

            fmt = xls.GetCellVisibleFormatDef(48, 4);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(48, 4, xls.AddFormat(fmt));
            xls.SetCellValue(48, 4, "Año 2");

            fmt = xls.GetCellVisibleFormatDef(48, 5);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(48, 5, xls.AddFormat(fmt));
            xls.SetCellValue(48, 5, "Año 3");

            fmt = xls.GetCellVisibleFormatDef(48, 6);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(48, 6, xls.AddFormat(fmt));
            xls.SetCellValue(48, 6, "Año 4");

            fmt = xls.GetCellVisibleFormatDef(48, 7);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(48, 7, xls.AddFormat(fmt));
            xls.SetCellValue(48, 7, "Año 5");

            fmt = xls.GetCellVisibleFormatDef(48, 8);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(48, 8, xls.AddFormat(fmt));
            xls.SetCellValue(48, 8, "Año 6");

            fmt = xls.GetCellVisibleFormatDef(48, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(48, 9, xls.AddFormat(fmt));
            xls.SetCellValue(48, 9, "Año 7");

            fmt = xls.GetCellVisibleFormatDef(48, 10);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(48, 10, xls.AddFormat(fmt));
            xls.SetCellValue(48, 10, "Año 8");

            fmt = xls.GetCellVisibleFormatDef(49, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(49, 1, xls.AddFormat(fmt));
            xls.SetCellValue(49, 1, "Desyerbe para mantenimiento");

            fmt = xls.GetCellVisibleFormatDef(49, 2);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(49, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(49, 3);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(49, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(49, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(49, 4, xls.AddFormat(fmt));
            xls.SetCellValue(49, 4, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!D50*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(49, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(49, 5, xls.AddFormat(fmt));
            xls.SetCellValue(49, 5, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!E50*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(49, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(49, 6, xls.AddFormat(fmt));
            xls.SetCellValue(49, 6, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!F50*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(49, 7);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(49, 7, xls.AddFormat(fmt));
            xls.SetCellValue(49, 7, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!G50*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(49, 8);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(49, 8, xls.AddFormat(fmt));
            xls.SetCellValue(49, 8, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!H50*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(49, 9);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(49, 9, xls.AddFormat(fmt));
            xls.SetCellValue(49, 9, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!I50*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(49, 10);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(49, 10, xls.AddFormat(fmt));
            xls.SetCellValue(49, 10, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!J50*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(50, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(50, 1, xls.AddFormat(fmt));
            xls.SetCellValue(50, 1, "Desyerbe quimico para mantenimiento");

            fmt = xls.GetCellVisibleFormatDef(50, 2);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(50, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(50, 3);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(50, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(50, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(50, 4, xls.AddFormat(fmt));
            xls.SetCellValue(50, 4, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!D51*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(50, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(50, 5, xls.AddFormat(fmt));
            xls.SetCellValue(50, 5, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!E51*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(50, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(50, 6, xls.AddFormat(fmt));
            xls.SetCellValue(50, 6, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!F51*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(50, 7);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(50, 7, xls.AddFormat(fmt));
            xls.SetCellValue(50, 7, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!G51*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(50, 8);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(50, 8, xls.AddFormat(fmt));
            xls.SetCellValue(50, 8, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!H51*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(50, 9);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(50, 9, xls.AddFormat(fmt));
            xls.SetCellValue(50, 9, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!I51*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(50, 10);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(50, 10, xls.AddFormat(fmt));
            xls.SetCellValue(50, 10, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!J51*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(51, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(51, 1, xls.AddFormat(fmt));
            xls.SetCellValue(51, 1, "Aplicación de abonos orgánicos para mantenimiento");

            fmt = xls.GetCellVisibleFormatDef(51, 2);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(51, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(51, 3);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(51, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(51, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(51, 4, xls.AddFormat(fmt));
            xls.SetCellValue(51, 4, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!D52*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(51, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(51, 5, xls.AddFormat(fmt));
            xls.SetCellValue(51, 5, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!E52*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(51, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(51, 6, xls.AddFormat(fmt));
            xls.SetCellValue(51, 6, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!F52*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(51, 7);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(51, 7, xls.AddFormat(fmt));
            xls.SetCellValue(51, 7, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!G52*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(51, 8);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(51, 8, xls.AddFormat(fmt));
            xls.SetCellValue(51, 8, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!H52*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(51, 9);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(51, 9, xls.AddFormat(fmt));
            xls.SetCellValue(51, 9, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!I52*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(51, 10);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(51, 10, xls.AddFormat(fmt));
            xls.SetCellValue(51, 10, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!J52*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(52, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(52, 1, xls.AddFormat(fmt));
            xls.SetCellValue(52, 1, "Aplicación de abonos químicos para mantenimiento");

            fmt = xls.GetCellVisibleFormatDef(52, 2);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(52, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(52, 3);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(52, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(52, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(52, 4, xls.AddFormat(fmt));
            xls.SetCellValue(52, 4, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!D53*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(52, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(52, 5, xls.AddFormat(fmt));
            xls.SetCellValue(52, 5, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!E53*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(52, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(52, 6, xls.AddFormat(fmt));
            xls.SetCellValue(52, 6, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!F53*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(52, 7);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(52, 7, xls.AddFormat(fmt));
            xls.SetCellValue(52, 7, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!G53*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(52, 8);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(52, 8, xls.AddFormat(fmt));
            xls.SetCellValue(52, 8, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!H53*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(52, 9);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(52, 9, xls.AddFormat(fmt));
            xls.SetCellValue(52, 9, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!I53*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(52, 10);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(52, 10, xls.AddFormat(fmt));
            xls.SetCellValue(52, 10, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!J53*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(53, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(53, 1, xls.AddFormat(fmt));
            xls.SetCellValue(53, 1, "Aplicación de foliares para fertilización y control roya");

            fmt = xls.GetCellVisibleFormatDef(53, 2);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(53, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(53, 3);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(53, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(53, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(53, 4, xls.AddFormat(fmt));
            xls.SetCellValue(53, 4, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!D54*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(53, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(53, 5, xls.AddFormat(fmt));
            xls.SetCellValue(53, 5, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!E54*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(53, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(53, 6, xls.AddFormat(fmt));
            xls.SetCellValue(53, 6, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!F54*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(53, 7);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(53, 7, xls.AddFormat(fmt));
            xls.SetCellValue(53, 7, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!G54*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(53, 8);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(53, 8, xls.AddFormat(fmt));
            xls.SetCellValue(53, 8, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!H54*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(53, 9);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(53, 9, xls.AddFormat(fmt));
            xls.SetCellValue(53, 9, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!I54*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(53, 10);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(53, 10, xls.AddFormat(fmt));
            xls.SetCellValue(53, 10, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!J54*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(54, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(54, 1, xls.AddFormat(fmt));
            xls.SetCellValue(54, 1, "Construcción de barreras vivas (rompe-vientos)");

            fmt = xls.GetCellVisibleFormatDef(54, 2);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(54, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(54, 3);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(54, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(54, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(54, 4, xls.AddFormat(fmt));
            xls.SetCellValue(54, 4, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!D55*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(54, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(54, 5, xls.AddFormat(fmt));
            xls.SetCellValue(54, 5, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!E55*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(54, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(54, 6, xls.AddFormat(fmt));
            xls.SetCellValue(54, 6, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!F55*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(54, 7);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(54, 7, xls.AddFormat(fmt));
            xls.SetCellValue(54, 7, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!G55*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(54, 8);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(54, 8, xls.AddFormat(fmt));
            xls.SetCellValue(54, 8, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!H55*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(54, 9);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(54, 9, xls.AddFormat(fmt));
            xls.SetCellValue(54, 9, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!I55*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(54, 10);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(54, 10, xls.AddFormat(fmt));
            xls.SetCellValue(54, 10, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!J55*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(55, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(55, 1, xls.AddFormat(fmt));
            xls.SetCellValue(55, 1, "Podas de árboles de sombra (sostenimiento)");

            fmt = xls.GetCellVisibleFormatDef(55, 2);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(55, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(55, 3);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(55, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(55, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(55, 4, xls.AddFormat(fmt));
            xls.SetCellValue(55, 4, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!D56*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(55, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(55, 5, xls.AddFormat(fmt));
            xls.SetCellValue(55, 5, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!E56*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(55, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(55, 6, xls.AddFormat(fmt));
            xls.SetCellValue(55, 6, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!F56*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(55, 7);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(55, 7, xls.AddFormat(fmt));
            xls.SetCellValue(55, 7, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!G56*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(55, 8);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(55, 8, xls.AddFormat(fmt));
            xls.SetCellValue(55, 8, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!H56*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(55, 9);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(55, 9, xls.AddFormat(fmt));
            xls.SetCellValue(55, 9, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!I56*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(55, 10);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(55, 10, xls.AddFormat(fmt));
            xls.SetCellValue(55, 10, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!J56*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(56, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.left;
            xls.SetCellFormat(56, 1, xls.AddFormat(fmt));
            xls.SetCellValue(56, 1, "Control de Broca (re-re, repela, fumigaciones)");

            fmt = xls.GetCellVisibleFormatDef(56, 2);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(56, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(56, 3);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(56, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(56, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(56, 4, xls.AddFormat(fmt));
            xls.SetCellValue(56, 4, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!D57*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(56, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(56, 5, xls.AddFormat(fmt));
            xls.SetCellValue(56, 5, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!E57*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(56, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(56, 6, xls.AddFormat(fmt));
            xls.SetCellValue(56, 6, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!F57*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(56, 7);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(56, 7, xls.AddFormat(fmt));
            xls.SetCellValue(56, 7, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!G57*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(56, 8);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(56, 8, xls.AddFormat(fmt));
            xls.SetCellValue(56, 8, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!H57*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(56, 9);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(56, 9, xls.AddFormat(fmt));
            xls.SetCellValue(56, 9, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!I57*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(56, 10);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(56, 10, xls.AddFormat(fmt));
            xls.SetCellValue(56, 10, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!J57*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(57, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(57, 1, xls.AddFormat(fmt));
            xls.SetCellValue(57, 1, "Manejo de tejido (desrrame o podas del café)");

            fmt = xls.GetCellVisibleFormatDef(57, 2);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(57, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(57, 3);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(57, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(57, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(57, 4, xls.AddFormat(fmt));
            xls.SetCellValue(57, 4, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!D58*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(57, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(57, 5, xls.AddFormat(fmt));
            xls.SetCellValue(57, 5, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!E58*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(57, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(57, 6, xls.AddFormat(fmt));
            xls.SetCellValue(57, 6, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!F58*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(57, 7);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(57, 7, xls.AddFormat(fmt));
            xls.SetCellValue(57, 7, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!G58*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(57, 8);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(57, 8, xls.AddFormat(fmt));
            xls.SetCellValue(57, 8, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!H58*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(57, 9);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(57, 9, xls.AddFormat(fmt));
            xls.SetCellValue(57, 9, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!I58*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(57, 10);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(57, 10, xls.AddFormat(fmt));
            xls.SetCellValue(57, 10, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!J58*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(58, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.left;
            xls.SetCellFormat(58, 1, xls.AddFormat(fmt));
            xls.SetCellValue(58, 1, "Otros:");

            fmt = xls.GetCellVisibleFormatDef(58, 2);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(58, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(58, 3);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(58, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(58, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(58, 4, xls.AddFormat(fmt));
            xls.SetCellValue(58, 4, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!D59*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(58, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(58, 5, xls.AddFormat(fmt));
            xls.SetCellValue(58, 5, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!E59*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(58, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(58, 6, xls.AddFormat(fmt));
            xls.SetCellValue(58, 6, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!F59*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(58, 7);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(58, 7, xls.AddFormat(fmt));
            xls.SetCellValue(58, 7, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!G59*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(58, 8);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(58, 8, xls.AddFormat(fmt));
            xls.SetCellValue(58, 8, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!H59*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(58, 9);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(58, 9, xls.AddFormat(fmt));
            xls.SetCellValue(58, 9, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!I59*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(58, 10);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(58, 10, xls.AddFormat(fmt));
            xls.SetCellValue(58, 10, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!J59*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(59, 1);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Font.Style = TFlxFontStyles.Italic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0x33, 0x66, 0xFF);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(59, 1, xls.AddFormat(fmt));
            xls.SetCellValue(59, 1, "Valor mano obra para mantenimiento");

            fmt = xls.GetCellVisibleFormatDef(59, 2);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(59, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(59, 3);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(59, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(59, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.149998474074526);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(59, 4, xls.AddFormat(fmt));
            xls.SetCellValue(59, 4, new TFormula("=SUM(D49:D58)"));

            fmt = xls.GetCellVisibleFormatDef(59, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.149998474074526);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(59, 5, xls.AddFormat(fmt));
            xls.SetCellValue(59, 5, new TFormula("=SUM(E49:E58)"));

            fmt = xls.GetCellVisibleFormatDef(59, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.149998474074526);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(59, 6, xls.AddFormat(fmt));
            xls.SetCellValue(59, 6, new TFormula("=SUM(F49:F58)"));

            fmt = xls.GetCellVisibleFormatDef(59, 7);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.149998474074526);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(59, 7, xls.AddFormat(fmt));
            xls.SetCellValue(59, 7, new TFormula("=SUM(G49:G58)"));

            fmt = xls.GetCellVisibleFormatDef(59, 8);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.149998474074526);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(59, 8, xls.AddFormat(fmt));
            xls.SetCellValue(59, 8, new TFormula("=SUM(H49:H58)"));

            fmt = xls.GetCellVisibleFormatDef(59, 9);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.149998474074526);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(59, 9, xls.AddFormat(fmt));
            xls.SetCellValue(59, 9, new TFormula("=SUM(I49:I58)"));

            fmt = xls.GetCellVisibleFormatDef(59, 10);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.149998474074526);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(59, 10, xls.AddFormat(fmt));
            xls.SetCellValue(59, 10, new TFormula("=SUM(J49:J58)"));

            fmt = xls.GetCellVisibleFormatDef(60, 1);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(60, 1, xls.AddFormat(fmt));
            xls.SetCellValue(60, 1, "Valor mano de obra cosecha");

            fmt = xls.GetCellVisibleFormatDef(60, 2);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(60, 2, xls.AddFormat(fmt));
            xls.SetCellValue(60, 2, "Año 0");

            fmt = xls.GetCellVisibleFormatDef(60, 3);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(60, 3, xls.AddFormat(fmt));
            xls.SetCellValue(60, 3, "Año 1");

            fmt = xls.GetCellVisibleFormatDef(60, 4);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(60, 4, xls.AddFormat(fmt));
            xls.SetCellValue(60, 4, "Año 2");

            fmt = xls.GetCellVisibleFormatDef(60, 5);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(60, 5, xls.AddFormat(fmt));
            xls.SetCellValue(60, 5, "Año 3");

            fmt = xls.GetCellVisibleFormatDef(60, 6);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(60, 6, xls.AddFormat(fmt));
            xls.SetCellValue(60, 6, "Año 4");

            fmt = xls.GetCellVisibleFormatDef(60, 7);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(60, 7, xls.AddFormat(fmt));
            xls.SetCellValue(60, 7, "Año 5");

            fmt = xls.GetCellVisibleFormatDef(60, 8);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(60, 8, xls.AddFormat(fmt));
            xls.SetCellValue(60, 8, "Año 6");

            fmt = xls.GetCellVisibleFormatDef(60, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(60, 9, xls.AddFormat(fmt));
            xls.SetCellValue(60, 9, "Año 7");

            fmt = xls.GetCellVisibleFormatDef(60, 10);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(60, 10, xls.AddFormat(fmt));
            xls.SetCellValue(60, 10, "Año 8");

            fmt = xls.GetCellVisibleFormatDef(61, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Locked = false;
            fmt.WrapText = true;
            xls.SetCellFormat(61, 1, xls.AddFormat(fmt));
            xls.SetCellValue(61, 1, "Recoleccion de café");

            fmt = xls.GetCellVisibleFormatDef(61, 2);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(61, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(61, 3);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(61, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(61, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(61, 4, xls.AddFormat(fmt));
            xls.SetCellValue(61, 4, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!D62*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(61, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(61, 5, xls.AddFormat(fmt));
            xls.SetCellValue(61, 5, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!E62*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(61, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(61, 6, xls.AddFormat(fmt));
            xls.SetCellValue(61, 6, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!F62*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(61, 7);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(61, 7, xls.AddFormat(fmt));
            xls.SetCellValue(61, 7, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!G62*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(61, 8);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(61, 8, xls.AddFormat(fmt));
            xls.SetCellValue(61, 8, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!H62*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(61, 9);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(61, 9, xls.AddFormat(fmt));
            xls.SetCellValue(61, 9, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!I62*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(61, 10);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(61, 10, xls.AddFormat(fmt));
            xls.SetCellValue(61, 10, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!J62*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(62, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Locked = false;
            fmt.WrapText = true;
            xls.SetCellFormat(62, 1, xls.AddFormat(fmt));
            xls.SetCellValue(62, 1, "Zarandeo del cerezo o rebalze (separar granos afectados por broca, dañados)");

            fmt = xls.GetCellVisibleFormatDef(62, 2);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(62, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(62, 3);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(62, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(62, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(62, 4, xls.AddFormat(fmt));
            xls.SetCellValue(62, 4, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!D63*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(62, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(62, 5, xls.AddFormat(fmt));
            xls.SetCellValue(62, 5, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!E63*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(62, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(62, 6, xls.AddFormat(fmt));
            xls.SetCellValue(62, 6, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!F63*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(62, 7);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(62, 7, xls.AddFormat(fmt));
            xls.SetCellValue(62, 7, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!G63*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(62, 8);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(62, 8, xls.AddFormat(fmt));
            xls.SetCellValue(62, 8, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!H63*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(62, 9);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(62, 9, xls.AddFormat(fmt));
            xls.SetCellValue(62, 9, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!I63*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(62, 10);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(62, 10, xls.AddFormat(fmt));
            xls.SetCellValue(62, 10, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!J63*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(63, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Locked = false;
            fmt.WrapText = true;
            xls.SetCellFormat(63, 1, xls.AddFormat(fmt));
            xls.SetCellValue(63, 1, "Otros:");

            fmt = xls.GetCellVisibleFormatDef(63, 2);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(63, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(63, 3);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(63, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(63, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(63, 4, xls.AddFormat(fmt));
            xls.SetCellValue(63, 4, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!D64*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(63, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(63, 5, xls.AddFormat(fmt));
            xls.SetCellValue(63, 5, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!E64*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(63, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(63, 6, xls.AddFormat(fmt));
            xls.SetCellValue(63, 6, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!F64*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(63, 7);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(63, 7, xls.AddFormat(fmt));
            xls.SetCellValue(63, 7, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!G64*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(63, 8);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(63, 8, xls.AddFormat(fmt));
            xls.SetCellValue(63, 8, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!H64*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(63, 9);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(63, 9, xls.AddFormat(fmt));
            xls.SetCellValue(63, 9, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!I64*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(63, 10);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(63, 10, xls.AddFormat(fmt));
            xls.SetCellValue(63, 10, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!J64*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(64, 1);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Font.Style = TFlxFontStyles.Italic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0x33, 0x66, 0xFF);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(64, 1, xls.AddFormat(fmt));
            xls.SetCellValue(64, 1, "Valor mano obra para cosecha");

            fmt = xls.GetCellVisibleFormatDef(64, 2);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(64, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(64, 3);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(64, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(64, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.149998474074526);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(64, 4, xls.AddFormat(fmt));
            xls.SetCellValue(64, 4, new TFormula("=SUM(D61:D63)"));

            fmt = xls.GetCellVisibleFormatDef(64, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.149998474074526);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(64, 5, xls.AddFormat(fmt));
            xls.SetCellValue(64, 5, new TFormula("=SUM(E61:E63)"));

            fmt = xls.GetCellVisibleFormatDef(64, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.149998474074526);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(64, 6, xls.AddFormat(fmt));
            xls.SetCellValue(64, 6, new TFormula("=SUM(F61:F63)"));

            fmt = xls.GetCellVisibleFormatDef(64, 7);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.149998474074526);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(64, 7, xls.AddFormat(fmt));
            xls.SetCellValue(64, 7, new TFormula("=SUM(G61:G63)"));

            fmt = xls.GetCellVisibleFormatDef(64, 8);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.149998474074526);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(64, 8, xls.AddFormat(fmt));
            xls.SetCellValue(64, 8, new TFormula("=SUM(H61:H63)"));

            fmt = xls.GetCellVisibleFormatDef(64, 9);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.149998474074526);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(64, 9, xls.AddFormat(fmt));
            xls.SetCellValue(64, 9, new TFormula("=SUM(I61:I63)"));

            fmt = xls.GetCellVisibleFormatDef(64, 10);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.149998474074526);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(64, 10, xls.AddFormat(fmt));
            xls.SetCellValue(64, 10, new TFormula("=SUM(J61:J63)"));

            fmt = xls.GetCellVisibleFormatDef(65, 1);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(65, 1, xls.AddFormat(fmt));
            xls.SetCellValue(65, 1, "Valor mano de obra para beneficio");

            fmt = xls.GetCellVisibleFormatDef(65, 2);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(65, 2, xls.AddFormat(fmt));
            xls.SetCellValue(65, 2, "Año 0");

            fmt = xls.GetCellVisibleFormatDef(65, 3);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(65, 3, xls.AddFormat(fmt));
            xls.SetCellValue(65, 3, "Año 1");

            fmt = xls.GetCellVisibleFormatDef(65, 4);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(65, 4, xls.AddFormat(fmt));
            xls.SetCellValue(65, 4, "Año 2");

            fmt = xls.GetCellVisibleFormatDef(65, 5);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(65, 5, xls.AddFormat(fmt));
            xls.SetCellValue(65, 5, "Año 3");

            fmt = xls.GetCellVisibleFormatDef(65, 6);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(65, 6, xls.AddFormat(fmt));
            xls.SetCellValue(65, 6, "Año 4");

            fmt = xls.GetCellVisibleFormatDef(65, 7);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(65, 7, xls.AddFormat(fmt));
            xls.SetCellValue(65, 7, "Año 5");

            fmt = xls.GetCellVisibleFormatDef(65, 8);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(65, 8, xls.AddFormat(fmt));
            xls.SetCellValue(65, 8, "Año 6");

            fmt = xls.GetCellVisibleFormatDef(65, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(65, 9, xls.AddFormat(fmt));
            xls.SetCellValue(65, 9, "Año 7");

            fmt = xls.GetCellVisibleFormatDef(65, 10);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(65, 10, xls.AddFormat(fmt));
            xls.SetCellValue(65, 10, "Año 8");

            fmt = xls.GetCellVisibleFormatDef(66, 1);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xCC, 0xFF, 0xCC);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(66, 1, xls.AddFormat(fmt));
            xls.SetCellValue(66, 1, "Beneficio humedo ");

            fmt = xls.GetCellVisibleFormatDef(66, 2);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(66, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(66, 3);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(66, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(66, 4);
            fmt.Format = "#,##0";
            xls.SetCellFormat(66, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(66, 5);
            fmt.Format = "#,##0";
            xls.SetCellFormat(66, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(66, 6);
            fmt.Format = "#,##0";
            xls.SetCellFormat(66, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(66, 7);
            fmt.Format = "#,##0";
            xls.SetCellFormat(66, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(66, 8);
            fmt.Format = "#,##0";
            xls.SetCellFormat(66, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(66, 9);
            fmt.Format = "#,##0";
            xls.SetCellFormat(66, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(66, 10);
            fmt.Format = "#,##0";
            xls.SetCellFormat(66, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(67, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(67, 1, xls.AddFormat(fmt));
            xls.SetCellValue(67, 1, "Despulpado y Fermentado");

            fmt = xls.GetCellVisibleFormatDef(67, 2);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(67, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(67, 3);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(67, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(67, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x80, 0x00);
            fmt.Format = "#,##0";
            xls.SetCellFormat(67, 4, xls.AddFormat(fmt));
            xls.SetCellValue(67, 4, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!D68*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(67, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x80, 0x00);
            fmt.Format = "#,##0";
            xls.SetCellFormat(67, 5, xls.AddFormat(fmt));
            xls.SetCellValue(67, 5, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!E68*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(67, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x80, 0x00);
            fmt.Format = "#,##0";
            xls.SetCellFormat(67, 6, xls.AddFormat(fmt));
            xls.SetCellValue(67, 6, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!F68*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(67, 7);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x80, 0x00);
            fmt.Format = "#,##0";
            xls.SetCellFormat(67, 7, xls.AddFormat(fmt));
            xls.SetCellValue(67, 7, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!G68*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(67, 8);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x80, 0x00);
            fmt.Format = "#,##0";
            xls.SetCellFormat(67, 8, xls.AddFormat(fmt));
            xls.SetCellValue(67, 8, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!H68*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(67, 9);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x80, 0x00);
            fmt.Format = "#,##0";
            xls.SetCellFormat(67, 9, xls.AddFormat(fmt));
            xls.SetCellValue(67, 9, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!I68*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(67, 10);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x80, 0x00);
            fmt.Format = "#,##0";
            xls.SetCellFormat(67, 10, xls.AddFormat(fmt));
            xls.SetCellValue(67, 10, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!J68*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(68, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(68, 1, xls.AddFormat(fmt));
            xls.SetCellValue(68, 1, "Lavado");

            fmt = xls.GetCellVisibleFormatDef(68, 2);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(68, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(68, 3);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(68, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(68, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x80, 0x00);
            fmt.Format = "#,##0";
            xls.SetCellFormat(68, 4, xls.AddFormat(fmt));
            xls.SetCellValue(68, 4, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!D69*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(68, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x80, 0x00);
            fmt.Format = "#,##0";
            xls.SetCellFormat(68, 5, xls.AddFormat(fmt));
            xls.SetCellValue(68, 5, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!E69*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(68, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x80, 0x00);
            fmt.Format = "#,##0";
            xls.SetCellFormat(68, 6, xls.AddFormat(fmt));
            xls.SetCellValue(68, 6, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!F69*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(68, 7);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x80, 0x00);
            fmt.Format = "#,##0";
            xls.SetCellFormat(68, 7, xls.AddFormat(fmt));
            xls.SetCellValue(68, 7, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!G69*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(68, 8);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x80, 0x00);
            fmt.Format = "#,##0";
            xls.SetCellFormat(68, 8, xls.AddFormat(fmt));
            xls.SetCellValue(68, 8, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!H69*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(68, 9);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x80, 0x00);
            fmt.Format = "#,##0";
            xls.SetCellFormat(68, 9, xls.AddFormat(fmt));
            xls.SetCellValue(68, 9, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!I69*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(68, 10);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x80, 0x00);
            fmt.Format = "#,##0";
            xls.SetCellFormat(68, 10, xls.AddFormat(fmt));
            xls.SetCellValue(68, 10, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!J69*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(69, 1);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Font.Style = TFlxFontStyles.Italic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0x33, 0x66, 0xFF);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(69, 1, xls.AddFormat(fmt));
            xls.SetCellValue(69, 1, "Valor mano obra para beneficio  humedo");

            fmt = xls.GetCellVisibleFormatDef(69, 2);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(69, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(69, 3);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(69, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(69, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.149998474074526);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(69, 4, xls.AddFormat(fmt));
            xls.SetCellValue(69, 4, new TFormula("=SUM(D67:D68)"));

            fmt = xls.GetCellVisibleFormatDef(69, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.149998474074526);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(69, 5, xls.AddFormat(fmt));
            xls.SetCellValue(69, 5, new TFormula("=SUM(E67:E68)"));

            fmt = xls.GetCellVisibleFormatDef(69, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.149998474074526);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(69, 6, xls.AddFormat(fmt));
            xls.SetCellValue(69, 6, new TFormula("=SUM(F67:F68)"));

            fmt = xls.GetCellVisibleFormatDef(69, 7);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.149998474074526);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(69, 7, xls.AddFormat(fmt));
            xls.SetCellValue(69, 7, new TFormula("=SUM(G67:G68)"));

            fmt = xls.GetCellVisibleFormatDef(69, 8);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.149998474074526);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(69, 8, xls.AddFormat(fmt));
            xls.SetCellValue(69, 8, new TFormula("=SUM(H67:H68)"));

            fmt = xls.GetCellVisibleFormatDef(69, 9);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.149998474074526);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(69, 9, xls.AddFormat(fmt));
            xls.SetCellValue(69, 9, new TFormula("=SUM(I67:I68)"));

            fmt = xls.GetCellVisibleFormatDef(69, 10);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.149998474074526);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(69, 10, xls.AddFormat(fmt));
            xls.SetCellValue(69, 10, new TFormula("=SUM(J67:J68)"));

            fmt = xls.GetCellVisibleFormatDef(70, 1);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xCC, 0xFF, 0xCC);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(70, 1, xls.AddFormat(fmt));
            xls.SetCellValue(70, 1, "Beneficio seco");

            fmt = xls.GetCellVisibleFormatDef(70, 2);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(70, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(70, 3);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(70, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(70, 4);
            fmt.Format = "#,##0";
            xls.SetCellFormat(70, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(70, 5);
            fmt.Format = "#,##0";
            xls.SetCellFormat(70, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(70, 6);
            fmt.Format = "#,##0";
            xls.SetCellFormat(70, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(70, 7);
            fmt.Format = "#,##0";
            xls.SetCellFormat(70, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(70, 8);
            fmt.Format = "#,##0";
            xls.SetCellFormat(70, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(70, 9);
            fmt.Format = "#,##0";
            xls.SetCellFormat(70, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(70, 10);
            fmt.Format = "#,##0";
            xls.SetCellFormat(70, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(71, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(71, 1, xls.AddFormat(fmt));
            xls.SetCellValue(71, 1, "Secado");

            fmt = xls.GetCellVisibleFormatDef(71, 2);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(71, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(71, 3);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(71, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(71, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(71, 4, xls.AddFormat(fmt));
            xls.SetCellValue(71, 4, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!D71*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(71, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(71, 5, xls.AddFormat(fmt));
            xls.SetCellValue(71, 5, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!E71*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(71, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(71, 6, xls.AddFormat(fmt));
            xls.SetCellValue(71, 6, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!F71*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(71, 7);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(71, 7, xls.AddFormat(fmt));
            xls.SetCellValue(71, 7, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!G71*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(71, 8);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(71, 8, xls.AddFormat(fmt));
            xls.SetCellValue(71, 8, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!H71*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(71, 9);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(71, 9, xls.AddFormat(fmt));
            xls.SetCellValue(71, 9, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!I71*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(71, 10);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(71, 10, xls.AddFormat(fmt));
            xls.SetCellValue(71, 10, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!J71*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(72, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(72, 1, xls.AddFormat(fmt));
            xls.SetCellValue(72, 1, "Zarandeo");

            fmt = xls.GetCellVisibleFormatDef(72, 2);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(72, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(72, 3);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(72, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(72, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(72, 4, xls.AddFormat(fmt));
            xls.SetCellValue(72, 4, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!D72*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(72, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(72, 5, xls.AddFormat(fmt));
            xls.SetCellValue(72, 5, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!E72*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(72, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(72, 6, xls.AddFormat(fmt));
            xls.SetCellValue(72, 6, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!F72*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(72, 7);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(72, 7, xls.AddFormat(fmt));
            xls.SetCellValue(72, 7, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!G72*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(72, 8);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(72, 8, xls.AddFormat(fmt));
            xls.SetCellValue(72, 8, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!H72*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(72, 9);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(72, 9, xls.AddFormat(fmt));
            xls.SetCellValue(72, 9, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!I72*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(72, 10);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(72, 10, xls.AddFormat(fmt));
            xls.SetCellValue(72, 10, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!J72*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(73, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(73, 1, xls.AddFormat(fmt));
            xls.SetCellValue(73, 1, "Escojo Selección");

            fmt = xls.GetCellVisibleFormatDef(73, 2);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(73, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(73, 3);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(73, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(73, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(73, 4, xls.AddFormat(fmt));
            xls.SetCellValue(73, 4, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!D73*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(73, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(73, 5, xls.AddFormat(fmt));
            xls.SetCellValue(73, 5, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!E73*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(73, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(73, 6, xls.AddFormat(fmt));
            xls.SetCellValue(73, 6, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!F73*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(73, 7);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(73, 7, xls.AddFormat(fmt));
            xls.SetCellValue(73, 7, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!G73*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(73, 8);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(73, 8, xls.AddFormat(fmt));
            xls.SetCellValue(73, 8, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!H73*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(73, 9);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(73, 9, xls.AddFormat(fmt));
            xls.SetCellValue(73, 9, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!I73*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(73, 10);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(73, 10, xls.AddFormat(fmt));
            xls.SetCellValue(73, 10, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!J73*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(74, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(74, 1, xls.AddFormat(fmt));
            xls.SetCellValue(74, 1, "Almacenamiento");

            fmt = xls.GetCellVisibleFormatDef(74, 2);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(74, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(74, 3);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(74, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(74, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(74, 4, xls.AddFormat(fmt));
            xls.SetCellValue(74, 4, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!D74*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(74, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(74, 5, xls.AddFormat(fmt));
            xls.SetCellValue(74, 5, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!E74*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(74, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(74, 6, xls.AddFormat(fmt));
            xls.SetCellValue(74, 6, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!F74*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(74, 7);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(74, 7, xls.AddFormat(fmt));
            xls.SetCellValue(74, 7, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!G74*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(74, 8);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(74, 8, xls.AddFormat(fmt));
            xls.SetCellValue(74, 8, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!H74*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(74, 9);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(74, 9, xls.AddFormat(fmt));
            xls.SetCellValue(74, 9, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!I74*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(74, 10);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(74, 10, xls.AddFormat(fmt));
            xls.SetCellValue(74, 10, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!J74*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(75, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(75, 1, xls.AddFormat(fmt));
            xls.SetCellValue(75, 1, "Aguas Miel");

            fmt = xls.GetCellVisibleFormatDef(75, 2);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(75, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(75, 3);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(75, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(75, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(75, 4, xls.AddFormat(fmt));
            xls.SetCellValue(75, 4, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!D75*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(75, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(75, 5, xls.AddFormat(fmt));
            xls.SetCellValue(75, 5, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!E75*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(75, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(75, 6, xls.AddFormat(fmt));
            xls.SetCellValue(75, 6, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!F75*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(75, 7);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(75, 7, xls.AddFormat(fmt));
            xls.SetCellValue(75, 7, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!G75*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(75, 8);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(75, 8, xls.AddFormat(fmt));
            xls.SetCellValue(75, 8, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!H75*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(75, 9);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(75, 9, xls.AddFormat(fmt));
            xls.SetCellValue(75, 9, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!I75*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(75, 10);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(75, 10, xls.AddFormat(fmt));
            xls.SetCellValue(75, 10, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!J75*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(76, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(76, 1, xls.AddFormat(fmt));
            xls.SetCellValue(76, 1, "Manejo de Pulpa");

            fmt = xls.GetCellVisibleFormatDef(76, 2);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(76, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(76, 3);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(76, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(76, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(76, 4, xls.AddFormat(fmt));
            xls.SetCellValue(76, 4, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!D76*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(76, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(76, 5, xls.AddFormat(fmt));
            xls.SetCellValue(76, 5, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!E76*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(76, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(76, 6, xls.AddFormat(fmt));
            xls.SetCellValue(76, 6, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!F76*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(76, 7);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(76, 7, xls.AddFormat(fmt));
            xls.SetCellValue(76, 7, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!G76*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(76, 8);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(76, 8, xls.AddFormat(fmt));
            xls.SetCellValue(76, 8, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!H76*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(76, 9);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(76, 9, xls.AddFormat(fmt));
            xls.SetCellValue(76, 9, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!I76*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(76, 10);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(76, 10, xls.AddFormat(fmt));
            xls.SetCellValue(76, 10, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!J76*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(77, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(77, 1, xls.AddFormat(fmt));
            xls.SetCellValue(77, 1, "Otros:");

            fmt = xls.GetCellVisibleFormatDef(77, 2);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(77, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(77, 3);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(77, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(77, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(77, 4, xls.AddFormat(fmt));
            xls.SetCellValue(77, 4, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!D77*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(77, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(77, 5, xls.AddFormat(fmt));
            xls.SetCellValue(77, 5, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!E77*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(77, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(77, 6, xls.AddFormat(fmt));
            xls.SetCellValue(77, 6, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!F77*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(77, 7);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(77, 7, xls.AddFormat(fmt));
            xls.SetCellValue(77, 7, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!G77*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(77, 8);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(77, 8, xls.AddFormat(fmt));
            xls.SetCellValue(77, 8, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!H77*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(77, 9);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(77, 9, xls.AddFormat(fmt));
            xls.SetCellValue(77, 9, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!I77*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(77, 10);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(77, 10, xls.AddFormat(fmt));
            xls.SetCellValue(77, 10, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!J77*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(78, 1);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Font.Style = TFlxFontStyles.Italic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0x33, 0x66, 0xFF);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(78, 1, xls.AddFormat(fmt));
            xls.SetCellValue(78, 1, "Valor mano obra para beneficio seco");

            fmt = xls.GetCellVisibleFormatDef(78, 2);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(78, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(78, 3);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(78, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(78, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.149998474074526);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(78, 4, xls.AddFormat(fmt));
            xls.SetCellValue(78, 4, new TFormula("=SUM(D71:D77)"));

            fmt = xls.GetCellVisibleFormatDef(78, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.149998474074526);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(78, 5, xls.AddFormat(fmt));
            xls.SetCellValue(78, 5, new TFormula("=SUM(E71:E77)"));

            fmt = xls.GetCellVisibleFormatDef(78, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.149998474074526);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(78, 6, xls.AddFormat(fmt));
            xls.SetCellValue(78, 6, new TFormula("=SUM(F71:F77)"));

            fmt = xls.GetCellVisibleFormatDef(78, 7);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.149998474074526);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(78, 7, xls.AddFormat(fmt));
            xls.SetCellValue(78, 7, new TFormula("=SUM(G71:G77)"));

            fmt = xls.GetCellVisibleFormatDef(78, 8);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.149998474074526);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(78, 8, xls.AddFormat(fmt));
            xls.SetCellValue(78, 8, new TFormula("=SUM(H71:H77)"));

            fmt = xls.GetCellVisibleFormatDef(78, 9);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.149998474074526);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(78, 9, xls.AddFormat(fmt));
            xls.SetCellValue(78, 9, new TFormula("=SUM(I71:I77)"));

            fmt = xls.GetCellVisibleFormatDef(78, 10);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.149998474074526);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(78, 10, xls.AddFormat(fmt));
            xls.SetCellValue(78, 10, new TFormula("=SUM(J71:J77)"));

            fmt = xls.GetCellVisibleFormatDef(79, 1);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(79, 1, xls.AddFormat(fmt));
            xls.SetCellValue(79, 1, "Valor mano de obra para cuestiones administrativas");

            fmt = xls.GetCellVisibleFormatDef(79, 2);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(79, 2, xls.AddFormat(fmt));
            xls.SetCellValue(79, 2, "Año 0");

            fmt = xls.GetCellVisibleFormatDef(79, 3);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(79, 3, xls.AddFormat(fmt));
            xls.SetCellValue(79, 3, "Año 1");

            fmt = xls.GetCellVisibleFormatDef(79, 4);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(79, 4, xls.AddFormat(fmt));
            xls.SetCellValue(79, 4, "Año 2");

            fmt = xls.GetCellVisibleFormatDef(79, 5);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(79, 5, xls.AddFormat(fmt));
            xls.SetCellValue(79, 5, "Año 3");

            fmt = xls.GetCellVisibleFormatDef(79, 6);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(79, 6, xls.AddFormat(fmt));
            xls.SetCellValue(79, 6, "Año 4");

            fmt = xls.GetCellVisibleFormatDef(79, 7);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(79, 7, xls.AddFormat(fmt));
            xls.SetCellValue(79, 7, "Año 5");

            fmt = xls.GetCellVisibleFormatDef(79, 8);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(79, 8, xls.AddFormat(fmt));
            xls.SetCellValue(79, 8, "Año 6");

            fmt = xls.GetCellVisibleFormatDef(79, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(79, 9, xls.AddFormat(fmt));
            xls.SetCellValue(79, 9, "Año 7");

            fmt = xls.GetCellVisibleFormatDef(79, 10);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(79, 10, xls.AddFormat(fmt));
            xls.SetCellValue(79, 10, "Año 8");

            fmt = xls.GetCellVisibleFormatDef(80, 1);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(80, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(80, 2);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(80, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(80, 3);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(80, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(80, 4);
            fmt.Format = "#,##0";
            xls.SetCellFormat(80, 4, xls.AddFormat(fmt));
            xls.SetCellValue(80, 4, new TFormula("=D69+D78"));

            fmt = xls.GetCellVisibleFormatDef(80, 5);
            fmt.Format = "#,##0";
            xls.SetCellFormat(80, 5, xls.AddFormat(fmt));
            xls.SetCellValue(80, 5, new TFormula("=E69+E78"));

            fmt = xls.GetCellVisibleFormatDef(80, 6);
            fmt.Format = "#,##0";
            xls.SetCellFormat(80, 6, xls.AddFormat(fmt));
            xls.SetCellValue(80, 6, new TFormula("=F69+F78"));

            fmt = xls.GetCellVisibleFormatDef(80, 7);
            fmt.Format = "#,##0";
            xls.SetCellFormat(80, 7, xls.AddFormat(fmt));
            xls.SetCellValue(80, 7, new TFormula("=G69+G78"));

            fmt = xls.GetCellVisibleFormatDef(80, 8);
            fmt.Format = "#,##0";
            xls.SetCellFormat(80, 8, xls.AddFormat(fmt));
            xls.SetCellValue(80, 8, new TFormula("=H69+H78"));

            fmt = xls.GetCellVisibleFormatDef(80, 9);
            fmt.Format = "#,##0";
            xls.SetCellFormat(80, 9, xls.AddFormat(fmt));
            xls.SetCellValue(80, 9, new TFormula("=I69+I78"));

            fmt = xls.GetCellVisibleFormatDef(80, 10);
            fmt.Format = "#,##0";
            xls.SetCellFormat(80, 10, xls.AddFormat(fmt));
            xls.SetCellValue(80, 10, new TFormula("=J69+J78"));

            fmt = xls.GetCellVisibleFormatDef(81, 1);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(81, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(81, 2);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(81, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(81, 3);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(81, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(81, 4);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(81, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(81, 5);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(81, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(81, 6);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(81, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(81, 7);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(81, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(81, 8);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(81, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(81, 9);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(81, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(81, 10);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(81, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(83, 1);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(83, 1, xls.AddFormat(fmt));
            xls.SetCellValue(83, 1, "Administración de su finca");

            fmt = xls.GetCellVisibleFormatDef(83, 2);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(83, 2, xls.AddFormat(fmt));
            xls.SetCellValue(83, 2, "Año 0");

            fmt = xls.GetCellVisibleFormatDef(83, 3);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(83, 3, xls.AddFormat(fmt));
            xls.SetCellValue(83, 3, "Año 1");

            fmt = xls.GetCellVisibleFormatDef(83, 4);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(83, 4, xls.AddFormat(fmt));
            xls.SetCellValue(83, 4, "Año 2");

            fmt = xls.GetCellVisibleFormatDef(83, 5);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(83, 5, xls.AddFormat(fmt));
            xls.SetCellValue(83, 5, "Año 3");

            fmt = xls.GetCellVisibleFormatDef(83, 6);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(83, 6, xls.AddFormat(fmt));
            xls.SetCellValue(83, 6, "Año 4");

            fmt = xls.GetCellVisibleFormatDef(83, 7);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(83, 7, xls.AddFormat(fmt));
            xls.SetCellValue(83, 7, "Año 5");

            fmt = xls.GetCellVisibleFormatDef(83, 8);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(83, 8, xls.AddFormat(fmt));
            xls.SetCellValue(83, 8, "Año 6");

            fmt = xls.GetCellVisibleFormatDef(83, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(83, 9, xls.AddFormat(fmt));
            xls.SetCellValue(83, 9, "Año 7");

            fmt = xls.GetCellVisibleFormatDef(83, 10);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(83, 10, xls.AddFormat(fmt));
            xls.SetCellValue(83, 10, "Año 8");

            fmt = xls.GetCellVisibleFormatDef(84, 1);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(84, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(84, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xBF, 0xBF, 0xBF);
            fmt.FillPattern.BgColor = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(84, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(84, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xBF, 0xBF, 0xBF);
            fmt.FillPattern.BgColor = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(84, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(84, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xBF, 0xBF, 0xBF);
            fmt.FillPattern.BgColor = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(84, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(84, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xBF, 0xBF, 0xBF);
            fmt.FillPattern.BgColor = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(84, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(84, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xBF, 0xBF, 0xBF);
            fmt.FillPattern.BgColor = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(84, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(84, 7);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xBF, 0xBF, 0xBF);
            fmt.FillPattern.BgColor = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(84, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(84, 8);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xBF, 0xBF, 0xBF);
            fmt.FillPattern.BgColor = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(84, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(84, 9);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xBF, 0xBF, 0xBF);
            fmt.FillPattern.BgColor = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(84, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(84, 10);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xBF, 0xBF, 0xBF);
            fmt.FillPattern.BgColor = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(84, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(85, 1);
            fmt.WrapText = true;
            xls.SetCellFormat(85, 1, xls.AddFormat(fmt));
            xls.SetCellValue(85, 1, "Cuantos dias al mes gasta ud en cuestiones administrativas de su finca como llevar"
            + " las cuentas, pagar servicios etc.?");

            fmt = xls.GetCellVisibleFormatDef(85, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x33, 0x66, 0xFF);
            xls.SetCellFormat(85, 2, xls.AddFormat(fmt));
            xls.SetCellValue(85, 2, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!B87*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(85, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x33, 0x66, 0xFF);
            xls.SetCellFormat(85, 3, xls.AddFormat(fmt));
            xls.SetCellValue(85, 3, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!C87*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(85, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x33, 0x66, 0xFF);
            xls.SetCellFormat(85, 4, xls.AddFormat(fmt));
            xls.SetCellValue(85, 4, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!D87*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(85, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x33, 0x66, 0xFF);
            xls.SetCellFormat(85, 5, xls.AddFormat(fmt));
            xls.SetCellValue(85, 5, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!E87*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(85, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x33, 0x66, 0xFF);
            xls.SetCellFormat(85, 6, xls.AddFormat(fmt));
            xls.SetCellValue(85, 6, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!F87*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(85, 7);
            fmt.Font.Color = TUIColor.FromArgb(0x33, 0x66, 0xFF);
            xls.SetCellFormat(85, 7, xls.AddFormat(fmt));
            xls.SetCellValue(85, 7, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!G87*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(85, 8);
            fmt.Font.Color = TUIColor.FromArgb(0x33, 0x66, 0xFF);
            xls.SetCellFormat(85, 8, xls.AddFormat(fmt));
            xls.SetCellValue(85, 8, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!H87*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(85, 9);
            fmt.Font.Color = TUIColor.FromArgb(0x33, 0x66, 0xFF);
            xls.SetCellFormat(85, 9, xls.AddFormat(fmt));
            xls.SetCellValue(85, 9, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!I87*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(85, 10);
            fmt.Font.Color = TUIColor.FromArgb(0x33, 0x66, 0xFF);
            xls.SetCellFormat(85, 10, xls.AddFormat(fmt));
            xls.SetCellValue(85, 10, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!J87*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(86, 1);
            fmt.WrapText = true;
            xls.SetCellFormat(86, 1, xls.AddFormat(fmt));
            xls.SetCellValue(86, 1, "Cuanto tiempo puede gastar Ud. Supervisando (no trabajando) actividades como limpias,"
            + " manejos, podas, obras conservación, cosecha etc");

            fmt = xls.GetCellVisibleFormatDef(86, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x33, 0x66, 0xFF);
            xls.SetCellFormat(86, 2, xls.AddFormat(fmt));
            xls.SetCellValue(86, 2, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!B85*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(86, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x33, 0x66, 0xFF);
            xls.SetCellFormat(86, 3, xls.AddFormat(fmt));
            xls.SetCellValue(86, 3, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!C85*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(86, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x33, 0x66, 0xFF);
            xls.SetCellFormat(86, 4, xls.AddFormat(fmt));
            xls.SetCellValue(86, 4, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!D85*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(86, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x33, 0x66, 0xFF);
            xls.SetCellFormat(86, 5, xls.AddFormat(fmt));
            xls.SetCellValue(86, 5, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!E85*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(86, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x33, 0x66, 0xFF);
            xls.SetCellFormat(86, 6, xls.AddFormat(fmt));
            xls.SetCellValue(86, 6, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!F85*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(86, 7);
            fmt.Font.Color = TUIColor.FromArgb(0x33, 0x66, 0xFF);
            xls.SetCellFormat(86, 7, xls.AddFormat(fmt));
            xls.SetCellValue(86, 7, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!G85*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(86, 8);
            fmt.Font.Color = TUIColor.FromArgb(0x33, 0x66, 0xFF);
            xls.SetCellFormat(86, 8, xls.AddFormat(fmt));
            xls.SetCellValue(86, 8, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!H85*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(86, 9);
            fmt.Font.Color = TUIColor.FromArgb(0x33, 0x66, 0xFF);
            xls.SetCellFormat(86, 9, xls.AddFormat(fmt));
            xls.SetCellValue(86, 9, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!I85*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(86, 10);
            fmt.Font.Color = TUIColor.FromArgb(0x33, 0x66, 0xFF);
            xls.SetCellFormat(86, 10, xls.AddFormat(fmt));
            xls.SetCellValue(86, 10, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!J85*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(87, 1);
            fmt.WrapText = true;
            xls.SetCellFormat(87, 1, xls.AddFormat(fmt));
            xls.SetCellValue(87, 1, "Cuanto tiempo puede gastar Ud. al año en capacitar a la gente que contrata para las"
            + " diversas labores de la finca");

            fmt = xls.GetCellVisibleFormatDef(87, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x33, 0x66, 0xFF);
            xls.SetCellFormat(87, 2, xls.AddFormat(fmt));
            xls.SetCellValue(87, 2, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!B89*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(87, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x33, 0x66, 0xFF);
            xls.SetCellFormat(87, 3, xls.AddFormat(fmt));
            xls.SetCellValue(87, 3, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!C89*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(87, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x33, 0x66, 0xFF);
            xls.SetCellFormat(87, 4, xls.AddFormat(fmt));
            xls.SetCellValue(87, 4, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!D89*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(87, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x33, 0x66, 0xFF);
            xls.SetCellFormat(87, 5, xls.AddFormat(fmt));
            xls.SetCellValue(87, 5, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!E89*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(87, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x33, 0x66, 0xFF);
            xls.SetCellFormat(87, 6, xls.AddFormat(fmt));
            xls.SetCellValue(87, 6, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!F89*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(87, 7);
            fmt.Font.Color = TUIColor.FromArgb(0x33, 0x66, 0xFF);
            xls.SetCellFormat(87, 7, xls.AddFormat(fmt));
            xls.SetCellValue(87, 7, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!G89*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(87, 8);
            fmt.Font.Color = TUIColor.FromArgb(0x33, 0x66, 0xFF);
            xls.SetCellFormat(87, 8, xls.AddFormat(fmt));
            xls.SetCellValue(87, 8, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!H89*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(87, 9);
            fmt.Font.Color = TUIColor.FromArgb(0x33, 0x66, 0xFF);
            xls.SetCellFormat(87, 9, xls.AddFormat(fmt));
            xls.SetCellValue(87, 9, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!I89*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(87, 10);
            fmt.Font.Color = TUIColor.FromArgb(0x33, 0x66, 0xFF);
            xls.SetCellFormat(87, 10, xls.AddFormat(fmt));
            xls.SetCellValue(87, 10, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!J89*'[Coffee Interactive"
            + " Tool 2.0  06_12_18.xlsx]Budget_Supuestos'!$B$71"));

            fmt = xls.GetCellVisibleFormatDef(88, 1);
            fmt.WrapText = true;
            xls.SetCellFormat(88, 1, xls.AddFormat(fmt));
            xls.SetCellValue(88, 1, "Cuanto puede gastar Ud. En costos extraordinarios tales como cubrir asistencias médicas"
            + " por accidentes de trabajo de sus trabajadores");

            fmt = xls.GetCellVisibleFormatDef(88, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x33, 0x66, 0xFF);
            xls.SetCellFormat(88, 2, xls.AddFormat(fmt));
            xls.SetCellValue(88, 2, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!B91"));

            fmt = xls.GetCellVisibleFormatDef(88, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x33, 0x66, 0xFF);
            xls.SetCellFormat(88, 3, xls.AddFormat(fmt));
            xls.SetCellValue(88, 3, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!C91"));

            fmt = xls.GetCellVisibleFormatDef(88, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x33, 0x66, 0xFF);
            xls.SetCellFormat(88, 4, xls.AddFormat(fmt));
            xls.SetCellValue(88, 4, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!D91"));

            fmt = xls.GetCellVisibleFormatDef(88, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x33, 0x66, 0xFF);
            xls.SetCellFormat(88, 5, xls.AddFormat(fmt));
            xls.SetCellValue(88, 5, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!E91"));

            fmt = xls.GetCellVisibleFormatDef(88, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x33, 0x66, 0xFF);
            xls.SetCellFormat(88, 6, xls.AddFormat(fmt));
            xls.SetCellValue(88, 6, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!F91"));

            fmt = xls.GetCellVisibleFormatDef(88, 7);
            fmt.Font.Color = TUIColor.FromArgb(0x33, 0x66, 0xFF);
            xls.SetCellFormat(88, 7, xls.AddFormat(fmt));
            xls.SetCellValue(88, 7, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!G91"));

            fmt = xls.GetCellVisibleFormatDef(88, 8);
            fmt.Font.Color = TUIColor.FromArgb(0x33, 0x66, 0xFF);
            xls.SetCellFormat(88, 8, xls.AddFormat(fmt));
            xls.SetCellValue(88, 8, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!H91"));

            fmt = xls.GetCellVisibleFormatDef(88, 9);
            fmt.Font.Color = TUIColor.FromArgb(0x33, 0x66, 0xFF);
            xls.SetCellFormat(88, 9, xls.AddFormat(fmt));
            xls.SetCellValue(88, 9, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!I91"));

            fmt = xls.GetCellVisibleFormatDef(88, 10);
            fmt.Font.Color = TUIColor.FromArgb(0x33, 0x66, 0xFF);
            xls.SetCellFormat(88, 10, xls.AddFormat(fmt));
            xls.SetCellValue(88, 10, new TFormula("='[Coffee Interactive Tool 2.0  06_12_18.xlsx]Budget_M Obra'!J91"));

            fmt = xls.GetCellVisibleFormatDef(89, 1);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.WrapText = true;
            xls.SetCellFormat(89, 1, xls.AddFormat(fmt));
            xls.SetCellValue(89, 1, "Total");
            xls.SetCellValue(89, 2, new TFormula("=SUM(B85:B88)"));
            xls.SetCellValue(89, 3, new TFormula("=SUM(C85:C88)"));
            xls.SetCellValue(89, 4, new TFormula("=SUM(D85:D88)"));
            xls.SetCellValue(89, 5, new TFormula("=SUM(E85:E88)"));
            xls.SetCellValue(89, 6, new TFormula("=SUM(F85:F88)"));
            xls.SetCellValue(89, 7, new TFormula("=SUM(G85:G88)"));
            xls.SetCellValue(89, 8, new TFormula("=SUM(H85:H88)"));
            xls.SetCellValue(89, 9, new TFormula("=SUM(I85:I88)"));
            xls.SetCellValue(89, 10, new TFormula("=SUM(J85:J88)"));

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
            xls.SetComment(61, 10, new TRichString("Juan Hernandez:\nIn 07/10/18 adjusted to multiply by the jornal price. Then dias de"
            + " recoleccion * jornal (no * cuanto paga por caja, este nos es el dato de numero de"
            + " cajas sino numero de dias) de recoleccion)", Runs, xls));

            //You probably don't need to call the lines below. This code is needed only if you want to change the comment box properties like color or default location
            TCommentProperties CommentProps = TCommentProperties.CreateStandard(61, 10, xls);
            CommentProps.Anchor = new TClientAnchor(TFlxAnchorType.DontMoveAndDontResize, 60, 134, 11, 213, 62, 18, 25, 142);

            //Excel by doesn't autofit the comment box so it can hold all text.
            //There is an option in TCommentProperties, but if you use it Excel will show the text in a single line.
            //To have FlexCel autofit the comment for you, you can do it with the following code:

            //    CommentProps.Anchor = xls.AutofitComment(new TRichString("Juan Hernandez:\nIn 07/10/18 adjusted to multiply by the jornal price. Then dias de"
            //    + " recoleccion * jornal (no * cuanto paga por caja, este nos es el dato de numero de"
            //    + " cajas sino numero de dias) de recoleccion)", Runs, xls), 1.5, true, 1.1, 0, CommentProps.Anchor);

            xls.SetCommentProperties(61, 10, CommentProps);

            //Cell selection and scroll position.
            xls.SelectCell(40, 13, false);

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
