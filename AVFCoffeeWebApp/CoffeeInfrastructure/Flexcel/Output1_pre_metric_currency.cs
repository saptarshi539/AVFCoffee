using CoffeeCore.Interfaces;
using FlexCel.Core;
using FlexCel.XlsAdapter;
using System;
using System.Collections.Generic;
using System.IO;
namespace CoffeeInfrastructure.Flexcel
{
    public class Output1_pre_metric_currency
    {

        public void Output1PreMetricCurrency(ExcelFile xls)
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

            xls.ActiveSheet = 35;    //Set the sheet we are working in.

            //Global Workbook Options
            xls.OptionsAutoCompressPictures = false;
            xls.OptionsMultithreadRecalc = 0;

            //Sheet Options
            xls.SheetName = "Outcome 1.0 pre_metric_currency";

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

            xls.SetColWidth(12, 12, 2528);    //(9.13 + 0.75) * 256
            xls.DefaultRowHeight = 315;

            xls.SetRowHeight(19, 375);    //18.75 * 20
            xls.SetRowHeight(21, 630);    //31.50 * 20
            xls.SetRowHeight(22, 630);    //31.50 * 20

            //Set the cell values
            TFlxFormat fmt;
            fmt = xls.GetCellVisibleFormatDef(6, 15);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(6, 15, xls.AddFormat(fmt));
            xls.SetCellValue(6, 15, "(Note to Programmer)");

            fmt = xls.GetCellVisibleFormatDef(6, 16);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(6, 16, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(6, 17);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(6, 17, xls.AddFormat(fmt));
            xls.SetCellValue(8, 15, "Please add in the graph the blue line according to the following linked value");
            xls.SetCellValue(8, 23, new TFormula("='Outcome TOTAL_Adj'!$P$18"));
            xls.SetCellValue(10, 15, "Please add to the graph the red line according to the price of coffee per pound in"
            + " the ");
            xls.SetCellValue(11, 15, "stock market");
            xls.SetCellValue(11, 23, 1.34);
            xls.SetCellValue(14, 15, "Y axis: US/POUND");

            fmt = xls.GetCellVisibleFormatDef(19, 11);
            fmt.Font.Size20 = 280;
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(19, 11, xls.AddFormat(fmt));
            xls.SetCellValue(19, 11, "US/ht");

            fmt = xls.GetCellVisibleFormatDef(19, 12);
            fmt.Font.Size20 = 280;
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(19, 12, xls.AddFormat(fmt));
            xls.SetCellValue(19, 12, "Pesos/ht");

            fmt = xls.GetCellVisibleFormatDef(21, 4);
            fmt.Font.Size20 = 480;
            xls.SetCellFormat(21, 4, xls.AddFormat(fmt));
            xls.SetCellValue(21, 4, "Your variable cost of production is: ");

            fmt = xls.GetCellVisibleFormatDef(21, 5);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(21, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(21, 6);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(21, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(21, 7);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(21, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(21, 8);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(21, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(21, 9);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(21, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(21, 10);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(21, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(21, 11);
            fmt.Font.Size20 = 280;
            fmt.Format = "0";
            xls.SetCellFormat(21, 11, xls.AddFormat(fmt));
            xls.SetCellValue(21, 11, new TFormula("='Outcome TOTAL_Adj'!$P$13"));

            fmt = xls.GetCellVisibleFormatDef(21, 12);
            fmt.Font.Size20 = 280;
            fmt.Format = "0";
            xls.SetCellFormat(21, 12, xls.AddFormat(fmt));
            xls.SetCellValue(21, 12, new TFormula("='Outcome TOTAL_Adj'!$P$5"));

            fmt = xls.GetCellVisibleFormatDef(21, 17);
            fmt.Format = "0";
            xls.SetCellFormat(21, 17, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(22, 4);
            fmt.Font.Size20 = 480;
            xls.SetCellFormat(22, 4, xls.AddFormat(fmt));
            xls.SetCellValue(22, 4, "Your total cost of production is: ");

            fmt = xls.GetCellVisibleFormatDef(22, 5);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(22, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(22, 6);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(22, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(22, 7);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(22, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(22, 8);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(22, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(22, 9);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(22, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(22, 10);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(22, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(22, 11);
            fmt.Font.Size20 = 280;
            fmt.Format = "0";
            xls.SetCellFormat(22, 11, xls.AddFormat(fmt));
            xls.SetCellValue(22, 11, new TFormula("='Outcome TOTAL_Adj'!$P$16"));

            fmt = xls.GetCellVisibleFormatDef(22, 12);
            fmt.Font.Size20 = 280;
            fmt.Format = "0";
            xls.SetCellFormat(22, 12, xls.AddFormat(fmt));
            xls.SetCellValue(22, 12, new TFormula("='Outcome TOTAL_Adj'!$P$8"));

            fmt = xls.GetCellVisibleFormatDef(22, 17);
            fmt.Format = "0";
            xls.SetCellFormat(22, 17, xls.AddFormat(fmt));

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
            xls.SetComment(11, 23, new TRichString("Juan Hernandez:\nFeel free to attach this number to a reliable source that update"
            + " each day.\nFrom now I am taking from:\nhttp://markets.businessinsider.com/commodities/coffee-price", Runs, xls));

            //You probably don't need to call the lines below. This code is needed only if you want to change the comment box properties like color or default location
            TCommentProperties CommentProps = TCommentProperties.CreateStandard(11, 23, xls);
            CommentProps.Anchor = new TClientAnchor(TFlxAnchorType.DontMoveAndDontResize, 10, 49, 24, 202, 17, 0, 29, 260);

            //Excel by doesn't autofit the comment box so it can hold all text.
            //There is an option in TCommentProperties, but if you use it Excel will show the text in a single line.
            //To have FlexCel autofit the comment for you, you can do it with the following code:

            //    CommentProps.Anchor = xls.AutofitComment(new TRichString("Juan Hernandez:\nFeel free to attach this number to a reliable source that update"
            //    + " each day.\nFrom now I am taking from:\nhttp://markets.businessinsider.com/commodities/coffee-price", Runs, xls), 1.5, true, 1.1, 0, CommentProps.Anchor);

            xls.SetCommentProperties(11, 23, CommentProps);

            //Objects
            TShapeProperties ShapeOptions2 = new TShapeProperties();
            ShapeOptions2.Anchor = new TClientAnchor(TFlxAnchorType.MoveAndResize, 7, 36, 4, 620, 7, 61, 12, 531);
            ShapeOptions2.ShapeType = TShapeType.Line;
            ShapeOptions2.ObjectType = TObjectType.MicrosoftOfficeDrawing;
            ShapeOptions2.ShapeName = "Straight Connector 7";
            ShapeOptions2.ShapeThemeFont = new TShapeFont(TFontScheme.Minor, TDrawingColor.FromTheme(TThemeColor.Foreground1));
            ShapeOptions2.FlipV = true;
            ShapeOptions2.Print = true;
            ShapeOptions2.Visible = true;
            ShapeOptions2.IsConnector = true;
            ShapeOptions2.ShapeGeometry = "<a:shapeGeom xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"><a:prstGeom"
            + " prst=\"line\"><a:avLst /></a:prstGeom></a:shapeGeom>";
            ShapeOptions2.ShapeOptions.SetValue(TShapeOption.fAutoTextMargin, true);
            ShapeOptions2.ShapeOptions.SetValue(TShapeOption.fillColor, 12419407);
            ShapeOptions2.ShapeOptions.SetValue(TShapeOption.fillBackColor, 134217808);
            ShapeOptions2.ShapeOptions.SetValue(TShapeOption.lineColor, 12419407);
            ShapeOptions2.ShapeOptions.SetValue(TShapeOption.lineWidth, 25400);
            ShapeOptions2.ShapeOptions.SetValue(TShapeOption.shadowColor, 0);
            ShapeOptions2.ShapeOptions.SetValue(TShapeOption.wzName, "Straight Connector 7");
            xls.AddAutoShape(ShapeOptions2);

            TShapeProperties ShapeOptions3 = new TShapeProperties();
            ShapeOptions3.Anchor = new TClientAnchor(TFlxAnchorType.MoveAndResize, 9, 49, 4, 591, 9, 85, 12, 518);
            ShapeOptions3.ShapeType = TShapeType.Line;
            ShapeOptions3.ObjectType = TObjectType.MicrosoftOfficeDrawing;
            ShapeOptions3.ShapeName = "Straight Connector 10";
            ShapeOptions3.ShapeThemeFont = new TShapeFont(TFontScheme.Minor, TDrawingColor.FromTheme(TThemeColor.Foreground1));
            ShapeOptions3.FlipV = true;
            ShapeOptions3.Print = true;
            ShapeOptions3.Visible = true;
            ShapeOptions3.IsConnector = true;
            ShapeOptions3.ShapeGeometry = "<a:shapeGeom xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"><a:prstGeom"
            + " prst=\"line\"><a:avLst /></a:prstGeom></a:shapeGeom>";
            ShapeOptions3.ShapeOptions.SetValue(TShapeOption.fAutoTextMargin, true);
            ShapeOptions3.ShapeOptions.SetValue(TShapeOption.fillColor, 12419407);
            ShapeOptions3.ShapeOptions.SetValue(TShapeOption.fillBackColor, 134217808);
            ShapeOptions3.ShapeOptions.SetValue(TShapeOption.lineColor, 255);
            ShapeOptions3.ShapeOptions.SetValue(TShapeOption.lineWidth, 25400);
            ShapeOptions3.ShapeOptions.SetValue(TShapeOption.shadowColor, 0);
            ShapeOptions3.ShapeOptions.SetValue(TShapeOption.wzName, "Straight Connector 10");
            xls.AddAutoShape(ShapeOptions3);


            //Cell selection and scroll position.
            xls.SelectCell(25, 8, false);
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
            //    xls.DocumentProperties.SetStandardProperty(TPropertyId.CreateTimeDate, "2015-01-08T03:31:31Z");

        }

    }
}
