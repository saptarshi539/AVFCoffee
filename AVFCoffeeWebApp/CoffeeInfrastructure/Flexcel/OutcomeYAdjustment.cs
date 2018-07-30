
using FlexCel.Core;
using System.IO;

namespace CoffeeInfrastructure.Flexcel
{
    public class OutcomeYAdjustment
    {
        public void Outcome_Y_Adjustment(ExcelFile xls)
        {
            xls.NewFile(31, TExcelFileFormat.v2010);    //Create a new Excel file with 31 sheets.

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
            xls.SheetName = "Budget_Presupuesto";
            xls.ActiveSheet = 25;
            xls.SheetName = "Budget_Valor de M Obra";
            xls.ActiveSheet = 26;
            xls.SheetName = "Budget_Establecimiento";
            xls.ActiveSheet = 27;
            xls.SheetName = "Budget_Sostenemiento";
            xls.ActiveSheet = 28;
            xls.SheetName = "Outcome 1.0 pre_metric_currency";
            xls.ActiveSheet = 29;
            xls.SheetName = "Conversiones";
            xls.ActiveSheet = 30;
            xls.SheetName = "Proporción de productividad";
            xls.ActiveSheet = 31;
            xls.SheetName = "Inputs 1.0 (Ref)";

            xls.ActiveSheet = 18;    //Set the sheet we are working in.

            //Global Workbook Options
            xls.OptionsAutoCompressPictures = false;
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
            Range = new TXlsNamedRange(RangeName, 26, 32, "=Budget_Establecimiento!$A$3:$C$53");
            //You could also use: Range = new TXlsNamedRange(RangeName, 26, 26, 3, 1, 53, 3, 32);
            xls.SetNamedRange(Range);

            RangeName = TXlsNamedRange.GetInternalName(InternalNameRange.Print_Area);
            Range = new TXlsNamedRange(RangeName, 23, 32, "='Budget_M Obra'!$A$1:$K$86");
            //You could also use: Range = new TXlsNamedRange(RangeName, 23, 23, 1, 1, 86, 11, 32);
            xls.SetNamedRange(Range);

            RangeName = TXlsNamedRange.GetInternalName(InternalNameRange.Print_Area);
            Range = new TXlsNamedRange(RangeName, 24, 32, "=Budget_Presupuesto!$A$34:$J$46");
            //You could also use: Range = new TXlsNamedRange(RangeName, 24, 24, 34, 1, 46, 10, 32);
            xls.SetNamedRange(Range);

            RangeName = TXlsNamedRange.GetInternalName(InternalNameRange.Print_Area);
            Range = new TXlsNamedRange(RangeName, 27, 32, "=Budget_Sostenemiento!$A$1:$K$44");
            //You could also use: Range = new TXlsNamedRange(RangeName, 27, 27, 1, 1, 44, 11, 32);
            xls.SetNamedRange(Range);

            RangeName = TXlsNamedRange.GetInternalName(InternalNameRange.Print_Area);
            Range = new TXlsNamedRange(RangeName, 21, 32, "=Budget_Supuestos!$A$276:$G$297");
            //You could also use: Range = new TXlsNamedRange(RangeName, 21, 21, 276, 1, 297, 7, 32);
            xls.SetNamedRange(Range);

            RangeName = TXlsNamedRange.GetInternalName(InternalNameRange.Print_Area);
            Range = new TXlsNamedRange(RangeName, 25, 32, "='Budget_Valor de M Obra'!$A$2:$J$85");
            //You could also use: Range = new TXlsNamedRange(RangeName, 25, 25, 2, 1, 85, 10, 32);
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

            //Set up rows and columns
            xls.DefaultColWidth = 2773;

            xls.SetColWidth(1, 2, 2773);    //(10.08 + 0.75) * 256

            xls.SetColWidth(3, 3, 5674);    //(21.41 + 0.75) * 256

            xls.SetColWidth(4, 4, 6485);    //(24.58 + 0.75) * 256

            xls.SetColWidth(5, 5, 3712);    //(13.75 + 0.75) * 256

            xls.SetColWidth(6, 6, 3541);    //(13.08 + 0.75) * 256

            xls.SetColWidth(7, 7, 3754);    //(13.91 + 0.75) * 256

            xls.SetColWidth(8, 8, 4309);    //(16.08 + 0.75) * 256

            xls.SetColWidth(9, 9, 5418);    //(20.41 + 0.75) * 256

            xls.SetColWidth(10, 15, 2773);    //(10.08 + 0.75) * 256

            xls.SetColWidth(16, 16, 3754);    //(13.91 + 0.75) * 256

            xls.SetColWidth(17, 16384, 2773);    //(10.08 + 0.75) * 256

            xls.SetRowHeight(2, 960);    //48.00 * 20

            //Set the cell values
            TFlxFormat fmt;
            fmt = xls.GetCellVisibleFormatDef(2, 3);
            fmt.WrapText = true;
            xls.SetCellFormat(2, 3, xls.AddFormat(fmt));
            xls.SetCellValue(2, 3, "Productivity original scale (Quintales/ht)");
            xls.SetCellValue(2, 4, new TFormula("='Inputs 1.0 Conv. new values'!$M$15"));
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

            //Images
            //using (FileStream fs = new FileStream("imagename.png", FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            //{
            //    TImageProperties ImgProps = new TImageProperties();
            //    ImgProps.Anchor = new TClientAnchor(TFlxAnchorType.MoveAndDontResize, 16, 0, 3, 0, 44, 51, 8, 183);
            //    ImgProps.ShapeName = "Picture 5";
            //    ImgProps.AltText = "Screen Shot 2017-09-12 at 7.33.13 PM.png";
            //    xls.AddImage(fs, ImgProps);
            //}

            //Cell selection and scroll position.
            xls.SelectCell(7, 9, false);

            //Standard Document Properties - Most are only for xlsx files. In xls files FlexCel will only change the Creation Date and Modified Date.
            xls.DocumentProperties.SetStandardProperty(TPropertyId.Author, "Mary Kate");

            //You will normally not set LastSavedBy, since this is a new file.
            //If you don't set it, FlexCel will use the creator instead.
            //    xls.DocumentProperties.SetStandardProperty(TPropertyId.LastSavedBy, "Juan Hernandez");


            //You will normally not set CreateDateTime, since this is a new file and FlexCel will automatically use the current datetime.
            //But if you are editing a file and want to preserve the original creation date, you need to either set PreserveCreationDate to true:
            //    xls.DocumentProperties.PreserveCreationDate = true;
            //Or you can hardcode a creating date by setting it in UTC time, ISO8601 format:
            //    xls.DocumentProperties.SetStandardProperty(TPropertyId.CreateTimeDate, "2015-01-08T03:31:31Z");

        }
    }
}
