using CoffeeCore.Interfaces;
using FlexCel.Core;
using FlexCel.XlsAdapter;
using System;
using System.Collections.Generic;
using System.IO;
namespace CoffeeInfrastructure.Flexcel
{
    public class InAdvanced
    {
        public void Inputs_advanced(ExcelFile xls)
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

            xls.ActiveSheet = 8;    //Set the sheet we are working in.

            //Global Workbook Options
            xls.OptionsAutoCompressPictures = false;
            xls.OptionsMultithreadRecalc = 0;

            //Sheet Options
            xls.SheetName = "Inputs advanced";
            xls.SheetZoom = 50;
            xls.SheetView = new TSheetView(TSheetViewType.Normal, true, true, 50, 50, 0);

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
            xls.PrintXResolution = 600;
            xls.PrintYResolution = 600;
            xls.PrintOptions = TPrintOptions.Orientation;
            xls.PrintPaperSize = TPaperSize.Letter;

            //Set up rows and columns
            xls.DefaultColWidth = 2261;

            xls.SetColWidth(3, 3, 4992);    //(18.75 + 0.75) * 256

            xls.SetColWidth(4, 4, 5077);    //(19.08 + 0.75) * 256

            xls.SetColWidth(5, 5, 18602);    //(71.91 + 0.75) * 256

            xls.SetColWidth(6, 6, 4650);    //(17.41 + 0.75) * 256

            xls.SetColWidth(7, 7, 9856);    //(37.75 + 0.75) * 256

            xls.SetColWidth(8, 8, 9130);    //(34.91 + 0.75) * 256

            xls.SetColWidth(17, 17, 2730);    //(9.91 + 0.75) * 256

            xls.SetRowHeight(3, 1400);    //70.00 * 20
            xls.SetRowHeight(4, 1380);    //69.00 * 20
            xls.SetRowHeight(5, 1050);    //52.50 * 20
            xls.SetRowHeight(6, 360);    //18.00 * 20
            xls.SetRowHeight(8, 620);    //31.00 * 20
            xls.SetRowHeight(9, 620);    //31.00 * 20
            xls.SetRowHeight(10, 620);    //31.00 * 20
            xls.SetRowHeight(11, 620);    //31.00 * 20
            xls.SetRowHeight(12, 360);    //18.00 * 20
            xls.SetRowHeight(13, 620);    //31.00 * 20
            xls.SetRowHeight(14, 620);    //31.00 * 20
            xls.SetRowHeight(15, 620);    //31.00 * 20
            xls.SetRowHeight(16, 620);    //31.00 * 20
            xls.SetRowHeight(17, 620);    //31.00 * 20
            xls.SetRowHeight(18, 620);    //31.00 * 20
            xls.SetRowHeight(19, 620);    //31.00 * 20
            xls.SetRowHeight(20, 620);    //31.00 * 20
            xls.SetRowHeight(21, 620);    //31.00 * 20
            xls.SetRowHeight(22, 620);    //31.00 * 20
            xls.SetRowHeight(23, 1080);    //54.00 * 20
            xls.SetRowHeight(24, 620);    //31.00 * 20
            xls.SetRowHeight(25, 620);    //31.00 * 20
            xls.SetRowHeight(26, 620);    //31.00 * 20
            xls.SetRowHeight(27, 620);    //31.00 * 20
            xls.SetRowHeight(28, 620);    //31.00 * 20
            xls.SetRowHeight(29, 620);    //31.00 * 20
            xls.SetRowHeight(30, 620);    //31.00 * 20
            xls.SetRowHeight(31, 620);    //31.00 * 20
            xls.SetRowHeight(32, 620);    //31.00 * 20
            xls.SetRowHeight(33, 620);    //31.00 * 20
            xls.SetRowHeight(34, 630);    //31.50 * 20
            xls.SetRowHeight(35, 360);    //18.00 * 20
            xls.SetRowHeight(36, 620);    //31.00 * 20
            xls.SetRowHeight(37, 620);    //31.00 * 20
            xls.SetRowHeight(38, 620);    //31.00 * 20
            xls.SetRowHeight(39, 620);    //31.00 * 20
            xls.SetRowHeight(40, 630);    //31.50 * 20
            xls.SetRowHeight(41, 400);    //20.00 * 20
            xls.SetRowHeight(42, 700);    //35.00 * 20
            xls.SetRowHeight(43, 620);    //31.00 * 20
            xls.SetRowHeight(44, 620);    //31.00 * 20
            xls.SetRowHeight(45, 620);    //31.00 * 20
            xls.SetRowHeight(46, 620);    //31.00 * 20
            xls.SetRowHeight(47, 700);    //35.00 * 20
            xls.SetRowHeight(48, 620);    //31.00 * 20
            xls.SetRowHeight(50, 620);    //31.00 * 20
            xls.SetRowHeight(51, 700);    //35.00 * 20
            xls.SetRowHeight(52, 620);    //31.00 * 20
            xls.SetRowHeight(53, 620);    //31.00 * 20
            xls.SetRowHeight(54, 620);    //31.00 * 20
            xls.SetRowHeight(55, 630);    //31.50 * 20
            xls.SetRowHeight(56, 510);    //25.50 * 20
            xls.SetRowHeight(57, 620);    //31.00 * 20
            xls.SetRowHeight(58, 620);    //31.00 * 20
            xls.SetRowHeight(59, 620);    //31.00 * 20
            xls.SetRowHeight(60, 700);    //35.00 * 20
            xls.SetRowHeight(61, 630);    //31.50 * 20
            xls.SetRowHeight(62, 360);    //18.00 * 20
            xls.SetRowHeight(63, 620);    //31.00 * 20
            xls.SetRowHeight(64, 620);    //31.00 * 20
            xls.SetRowHeight(65, 620);    //31.00 * 20
            xls.SetRowHeight(66, 620);    //31.00 * 20
            xls.SetRowHeight(67, 620);    //31.00 * 20
            xls.SetRowHeight(68, 620);    //31.00 * 20
            xls.SetRowHeight(69, 620);    //31.00 * 20
            xls.SetRowHeight(70, 620);    //31.00 * 20
            xls.SetRowHeight(71, 620);    //31.00 * 20
            xls.SetRowHeight(72, 620);    //31.00 * 20
            xls.SetRowHeight(73, 360);    //18.00 * 20
            xls.SetRowHeight(74, 620);    //31.00 * 20
            xls.SetRowHeight(75, 620);    //31.00 * 20
            xls.SetRowHeight(76, 620);    //31.00 * 20
            xls.SetRowHeight(77, 620);    //31.00 * 20
            xls.SetRowHeight(78, 620);    //31.00 * 20
            xls.SetRowHeight(79, 360);    //18.00 * 20
            xls.SetRowHeight(80, 620);    //31.00 * 20
            xls.SetRowHeight(81, 620);    //31.00 * 20
            xls.SetRowHeight(82, 620);    //31.00 * 20
            xls.SetRowHeight(83, 620);    //31.00 * 20
            xls.SetRowHeight(84, 620);    //31.00 * 20
            xls.SetRowHeight(85, 620);    //31.00 * 20
            xls.SetRowHeight(86, 620);    //31.00 * 20
            xls.SetRowHeight(87, 620);    //31.00 * 20
            xls.SetRowHeight(88, 630);    //31.50 * 20
            xls.SetRowHeight(89, 360);    //18.00 * 20
            xls.SetRowHeight(90, 620);    //31.00 * 20
            xls.SetRowHeight(91, 620);    //31.00 * 20
            xls.SetRowHeight(92, 620);    //31.00 * 20
            xls.SetRowHeight(93, 620);    //31.00 * 20
            xls.SetRowHeight(94, 620);    //31.00 * 20
            xls.SetRowHeight(95, 620);    //31.00 * 20
            xls.SetRowHeight(96, 620);    //31.00 * 20
            xls.SetRowHeight(97, 620);    //31.00 * 20
            xls.SetRowHeight(98, 620);    //31.00 * 20
            xls.SetRowHeight(99, 620);    //31.00 * 20
            xls.SetRowHeight(100, 360);    //18.00 * 20
            xls.SetRowHeight(101, 620);    //31.00 * 20
            xls.SetRowHeight(102, 620);    //31.00 * 20
            xls.SetRowHeight(103, 620);    //31.00 * 20
            xls.SetRowHeight(104, 620);    //31.00 * 20
            xls.SetRowHeight(105, 620);    //31.00 * 20
            xls.SetRowHeight(106, 360);    //18.00 * 20
            xls.SetRowHeight(107, 620);    //31.00 * 20
            xls.SetRowHeight(108, 620);    //31.00 * 20
            xls.SetRowHeight(109, 620);    //31.00 * 20
            xls.SetRowHeight(110, 620);    //31.00 * 20
            xls.SetRowHeight(111, 620);    //31.00 * 20
            xls.SetRowHeight(112, 620);    //31.00 * 20
            xls.SetRowHeight(113, 620);    //31.00 * 20
            xls.SetRowHeight(114, 620);    //31.00 * 20
            xls.SetRowHeight(115, 630);    //31.50 * 20
            xls.SetRowHeight(116, 360);    //18.00 * 20
            xls.SetRowHeight(117, 620);    //31.00 * 20
            xls.SetRowHeight(118, 620);    //31.00 * 20
            xls.SetRowHeight(119, 620);    //31.00 * 20
            xls.SetRowHeight(120, 620);    //31.00 * 20
            xls.SetRowHeight(121, 620);    //31.00 * 20
            xls.SetRowHeight(122, 620);    //31.00 * 20
            xls.SetRowHeight(123, 620);    //31.00 * 20
            xls.SetRowHeight(124, 620);    //31.00 * 20
            xls.SetRowHeight(125, 620);    //31.00 * 20
            xls.SetRowHeight(126, 620);    //31.00 * 20
            xls.SetRowHeight(127, 360);    //18.00 * 20
            xls.SetRowHeight(128, 620);    //31.00 * 20
            xls.SetRowHeight(129, 620);    //31.00 * 20
            xls.SetRowHeight(130, 620);    //31.00 * 20
            xls.SetRowHeight(131, 620);    //31.00 * 20
            xls.SetRowHeight(132, 620);    //31.00 * 20
            xls.SetRowHeight(133, 360);    //18.00 * 20
            xls.SetRowHeight(134, 620);    //31.00 * 20
            xls.SetRowHeight(135, 620);    //31.00 * 20
            xls.SetRowHeight(136, 620);    //31.00 * 20
            xls.SetRowHeight(137, 620);    //31.00 * 20
            xls.SetRowHeight(138, 620);    //31.00 * 20
            xls.SetRowHeight(139, 620);    //31.00 * 20
            xls.SetRowHeight(140, 620);    //31.00 * 20
            xls.SetRowHeight(141, 620);    //31.00 * 20
            xls.SetRowHeight(142, 630);    //31.50 * 20
            xls.SetRowHeight(143, 620);    //31.00 * 20
            xls.SetRowHeight(144, 620);    //31.00 * 20
            xls.SetRowHeight(145, 400);    //20.00 * 20
            xls.SetRowHeight(146, 700);    //35.00 * 20
            xls.SetRowHeight(147, 620);    //31.00 * 20
            xls.SetRowHeight(148, 400);    //20.00 * 20
            xls.SetRowHeight(149, 620);    //31.00 * 20
            xls.SetRowHeight(150, 620);    //31.00 * 20
            xls.SetRowHeight(151, 620);    //31.00 * 20
            xls.SetRowHeight(152, 930);    //46.50 * 20
            xls.SetRowHeight(153, 630);    //31.50 * 20
            xls.SetRowHeight(154, 630);    //31.50 * 20
            xls.SetRowHeight(155, 630);    //31.50 * 20
            xls.SetRowHeight(156, 350);    //17.50 * 20
            xls.SetRowHeight(157, 620);    //31.00 * 20
            xls.SetRowHeight(158, 620);    //31.00 * 20
            xls.SetRowHeight(159, 620);    //31.00 * 20
            xls.SetRowHeight(160, 620);    //31.00 * 20
            xls.SetRowHeight(161, 620);    //31.00 * 20
            xls.SetRowHeight(162, 620);    //31.00 * 20
            xls.SetRowHeight(163, 620);    //31.00 * 20
            xls.SetRowHeight(164, 620);    //31.00 * 20
            xls.SetRowHeight(165, 620);    //31.00 * 20
            xls.SetRowHeight(166, 620);    //31.00 * 20
            xls.SetRowHeight(167, 620);    //31.00 * 20
            xls.SetRowHeight(168, 620);    //31.00 * 20
            xls.SetRowHeight(169, 620);    //31.00 * 20
            xls.SetRowHeight(170, 620);    //31.00 * 20
            xls.SetRowHeight(171, 620);    //31.00 * 20
            xls.SetRowHeight(172, 620);    //31.00 * 20
            xls.SetRowHeight(173, 620);    //31.00 * 20
            xls.SetRowHeight(174, 620);    //31.00 * 20
            xls.SetRowHeight(175, 620);    //31.00 * 20
            xls.SetRowHeight(176, 620);    //31.00 * 20
            xls.SetRowHeight(178, 620);    //31.00 * 20
            xls.SetRowHeight(179, 620);    //31.00 * 20
            xls.SetRowHeight(180, 620);    //31.00 * 20
            xls.SetRowHeight(182, 620);    //31.00 * 20
            xls.SetRowHeight(183, 620);    //31.00 * 20
            xls.SetRowHeight(184, 620);    //31.00 * 20
            xls.SetRowHeight(185, 620);    //31.00 * 20
            xls.SetRowHeight(186, 700);    //35.00 * 20
            xls.SetRowHeight(187, 700);    //35.00 * 20
            xls.SetRowHeight(189, 620);    //31.00 * 20
            xls.SetRowHeight(190, 620);    //31.00 * 20
            xls.SetRowHeight(191, 620);    //31.00 * 20
            xls.SetRowHeight(192, 620);    //31.00 * 20
            xls.SetRowHeight(193, 620);    //31.00 * 20
            xls.SetRowHeight(194, 620);    //31.00 * 20
            xls.SetRowHeight(195, 460);    //23.00 * 20
            xls.SetRowHeight(196, 720);    //36.00 * 20
            xls.SetRowHeight(197, 1050);    //52.50 * 20
            xls.SetRowHeight(198, 620);    //31.00 * 20
            xls.SetRowHeight(199, 620);    //31.00 * 20
            xls.SetRowHeight(200, 620);    //31.00 * 20
            xls.SetRowHeight(201, 620);    //31.00 * 20
            xls.SetRowHeight(202, 700);    //35.00 * 20
            xls.SetRowHeight(203, 700);    //35.00 * 20
            xls.SetRowHeight(204, 700);    //35.00 * 20
            xls.SetRowHeight(205, 620);    //31.00 * 20
            xls.SetRowHeight(206, 1050);    //52.50 * 20
            xls.SetRowHeight(207, 700);    //35.00 * 20
            xls.SetRowHeight(208, 700);    //35.00 * 20
            xls.SetRowHeight(209, 620);    //31.00 * 20
            xls.SetRowHeight(210, 620);    //31.00 * 20
            xls.SetRowHeight(211, 620);    //31.00 * 20
            xls.SetRowHeight(212, 620);    //31.00 * 20
            xls.SetRowHeight(213, 620);    //31.00 * 20
            xls.SetRowHeight(214, 620);    //31.00 * 20
            xls.SetRowHeight(215, 620);    //31.00 * 20
            xls.SetRowHeight(216, 620);    //31.00 * 20
            xls.SetRowHeight(217, 620);    //31.00 * 20
            xls.SetRowHeight(218, 460);    //23.00 * 20
            xls.SetRowHeight(219, 700);    //35.00 * 20
            xls.SetRowHeight(220, 620);    //31.00 * 20
            xls.SetRowHeight(221, 620);    //31.00 * 20
            xls.SetRowHeight(222, 620);    //31.00 * 20
            xls.SetRowHeight(223, 700);    //35.00 * 20
            xls.SetRowHeight(224, 620);    //31.00 * 20
            xls.SetRowHeight(225, 620);    //31.00 * 20
            xls.SetRowHeight(226, 350);    //17.50 * 20
            xls.SetRowHeight(227, 620);    //31.00 * 20
            xls.SetRowHeight(228, 620);    //31.00 * 20
            xls.SetRowHeight(229, 460);    //23.00 * 20
            xls.SetRowHeight(230, 620);    //31.00 * 20
            xls.SetRowHeight(231, 620);    //31.00 * 20
            xls.SetRowHeight(232, 620);    //31.00 * 20
            xls.SetRowHeight(233, 620);    //31.00 * 20
            xls.SetRowHeight(234, 620);    //31.00 * 20
            xls.SetRowHeight(235, 350);    //17.50 * 20
            xls.SetRowHeight(236, 620);    //31.00 * 20
            xls.SetRowHeight(237, 620);    //31.00 * 20
            xls.SetRowHeight(238, 620);    //31.00 * 20
            xls.SetRowHeight(239, 620);    //31.00 * 20
            xls.SetRowHeight(240, 620);    //31.00 * 20
            xls.SetRowHeight(241, 620);    //31.00 * 20
            xls.SetRowHeight(242, 620);    //31.00 * 20
            xls.SetRowHeight(243, 620);    //31.00 * 20
            xls.SetRowHeight(244, 620);    //31.00 * 20
            xls.SetRowHeight(245, 350);    //17.50 * 20
            xls.SetRowHeight(246, 620);    //31.00 * 20
            xls.SetRowHeight(247, 620);    //31.00 * 20
            xls.SetRowHeight(248, 620);    //31.00 * 20
            xls.SetRowHeight(249, 620);    //31.00 * 20
            xls.SetRowHeight(250, 500);    //25.00 * 20
            xls.SetRowHeight(251, 350);    //17.50 * 20
            xls.SetRowHeight(252, 620);    //31.00 * 20
            xls.SetRowHeight(253, 620);    //31.00 * 20
            xls.SetRowHeight(254, 620);    //31.00 * 20
            xls.SetRowHeight(255, 620);    //31.00 * 20
            xls.SetRowHeight(256, 620);    //31.00 * 20
            xls.SetRowHeight(257, 620);    //31.00 * 20
            xls.SetRowHeight(258, 620);    //31.00 * 20
            xls.SetRowHeight(259, 350);    //17.50 * 20
            xls.SetRowHeight(260, 620);    //31.00 * 20
            xls.SetRowHeight(261, 620);    //31.00 * 20
            xls.SetRowHeight(262, 620);    //31.00 * 20
            xls.SetRowHeight(263, 620);    //31.00 * 20
            xls.SetRowHeight(264, 620);    //31.00 * 20
            xls.SetRowHeight(265, 620);    //31.00 * 20
            xls.SetRowHeight(266, 620);    //31.00 * 20
            xls.SetRowHeight(267, 620);    //31.00 * 20
            xls.SetRowHeight(268, 620);    //31.00 * 20
            xls.SetRowHeight(269, 620);    //31.00 * 20
            xls.SetRowHeight(270, 350);    //17.50 * 20
            xls.SetRowHeight(271, 350);    //17.50 * 20
            xls.SetRowHeight(272, 620);    //31.00 * 20
            xls.SetRowHeight(273, 740);    //37.00 * 20
            xls.SetRowHeight(274, 630);    //31.50 * 20
            xls.SetRowHeight(275, 620);    //31.00 * 20
            xls.SetRowHeight(276, 620);    //31.00 * 20
            xls.SetRowHeight(278, 580);    //29.00 * 20
            xls.SetRowHeight(281, 740);    //37.00 * 20
            xls.SetRowHeight(282, 820);    //41.00 * 20
            xls.SetRowHeight(286, 600);    //30.00 * 20
            xls.SetRowHeight(289, 1020);    //51.00 * 20
            xls.SetRowHeight(290, 640);    //32.00 * 20
            xls.SetRowHeight(295, 780);    //39.00 * 20
            xls.SetRowHeight(297, 400);    //20.00 * 20
            xls.SetRowHeight(301, 400);    //20.00 * 20
            xls.SetRowHeight(303, 620);    //31.00 * 20
            xls.SetRowHeight(304, 620);    //31.00 * 20
            xls.SetRowHeight(305, 400);    //20.00 * 20
            xls.SetRowHeight(306, 620);    //31.00 * 20
            xls.SetRowHeight(307, 620);    //31.00 * 20
            xls.SetRowHeight(308, 620);    //31.00 * 20
            xls.SetRowHeight(309, 620);    //31.00 * 20
            xls.SetRowHeight(310, 620);    //31.00 * 20
            xls.SetRowHeight(311, 620);    //31.00 * 20
            xls.SetRowHeight(312, 620);    //31.00 * 20
            xls.SetRowHeight(313, 620);    //31.00 * 20
            xls.SetRowHeight(314, 620);    //31.00 * 20
            xls.SetRowHeight(315, 620);    //31.00 * 20
            xls.SetRowHeight(316, 620);    //31.00 * 20
            xls.SetRowHeight(317, 620);    //31.00 * 20
            xls.SetRowHeight(318, 620);    //31.00 * 20
            xls.SetRowHeight(319, 620);    //31.00 * 20
            xls.SetRowHeight(320, 620);    //31.00 * 20
            xls.SetRowHeight(321, 620);    //31.00 * 20
            xls.SetRowHeight(322, 620);    //31.00 * 20
            xls.SetRowHeight(323, 620);    //31.00 * 20
            xls.SetRowHeight(324, 620);    //31.00 * 20
            xls.SetRowHeight(325, 620);    //31.00 * 20
            xls.SetRowHeight(326, 620);    //31.00 * 20
            xls.SetRowHeight(327, 620);    //31.00 * 20
            xls.SetRowHeight(328, 620);    //31.00 * 20
            xls.SetRowHeight(329, 620);    //31.00 * 20
            xls.SetRowHeight(330, 620);    //31.00 * 20
            xls.SetRowHeight(331, 620);    //31.00 * 20
            xls.SetRowHeight(332, 620);    //31.00 * 20
            xls.SetRowHeight(333, 620);    //31.00 * 20
            xls.SetRowHeight(334, 620);    //31.00 * 20
            xls.SetRowHeight(335, 620);    //31.00 * 20
            xls.SetRowHeight(336, 350);    //17.50 * 20
            xls.SetRowHeight(337, 620);    //31.00 * 20
            xls.SetRowHeight(338, 620);    //31.00 * 20
            xls.SetRowHeight(339, 620);    //31.00 * 20
            xls.SetRowHeight(340, 620);    //31.00 * 20
            xls.SetRowHeight(341, 620);    //31.00 * 20
            xls.SetRowHeight(342, 620);    //31.00 * 20
            xls.SetRowHeight(343, 620);    //31.00 * 20
            xls.SetRowHeight(344, 620);    //31.00 * 20
            xls.SetRowHeight(345, 620);    //31.00 * 20
            xls.SetRowHeight(346, 620);    //31.00 * 20
            xls.SetRowHeight(347, 620);    //31.00 * 20
            xls.SetRowHeight(348, 620);    //31.00 * 20
            xls.SetRowHeight(349, 620);    //31.00 * 20
            xls.SetRowHeight(350, 620);    //31.00 * 20
            xls.SetRowHeight(351, 620);    //31.00 * 20
            xls.SetRowHeight(352, 620);    //31.00 * 20
            xls.SetRowHeight(353, 620);    //31.00 * 20
            xls.SetRowHeight(354, 620);    //31.00 * 20
            xls.SetRowHeight(355, 620);    //31.00 * 20
            xls.SetRowHeight(356, 620);    //31.00 * 20
            xls.SetRowHeight(357, 350);    //17.50 * 20
            xls.SetRowHeight(358, 620);    //31.00 * 20
            xls.SetRowHeight(359, 620);    //31.00 * 20
            xls.SetRowHeight(360, 620);    //31.00 * 20
            xls.SetRowHeight(361, 620);    //31.00 * 20
            xls.SetRowHeight(362, 620);    //31.00 * 20
            xls.SetRowHeight(363, 620);    //31.00 * 20
            xls.SetRowHeight(364, 620);    //31.00 * 20
            xls.SetRowHeight(365, 620);    //31.00 * 20
            xls.SetRowHeight(366, 620);    //31.00 * 20
            xls.SetRowHeight(367, 620);    //31.00 * 20
            xls.SetRowHeight(368, 620);    //31.00 * 20
            xls.SetRowHeight(369, 620);    //31.00 * 20
            xls.SetRowHeight(370, 620);    //31.00 * 20
            xls.SetRowHeight(371, 620);    //31.00 * 20
            xls.SetRowHeight(372, 620);    //31.00 * 20
            xls.SetRowHeight(373, 620);    //31.00 * 20
            xls.SetRowHeight(374, 620);    //31.00 * 20
            xls.SetRowHeight(375, 620);    //31.00 * 20
            xls.SetRowHeight(376, 620);    //31.00 * 20
            xls.SetRowHeight(377, 620);    //31.00 * 20
            xls.SetRowHeight(378, 620);    //31.00 * 20
            xls.SetRowHeight(379, 620);    //31.00 * 20
            xls.SetRowHeight(380, 620);    //31.00 * 20
            xls.SetRowHeight(381, 620);    //31.00 * 20
            xls.SetRowHeight(382, 620);    //31.00 * 20
            xls.SetRowHeight(383, 620);    //31.00 * 20
            xls.SetRowHeight(384, 620);    //31.00 * 20
            xls.SetRowHeight(385, 620);    //31.00 * 20
            xls.SetRowHeight(386, 620);    //31.00 * 20
            xls.SetRowHeight(387, 620);    //31.00 * 20
            xls.SetRowHeight(388, 620);    //31.00 * 20
            xls.SetRowHeight(389, 620);    //31.00 * 20
            xls.SetRowHeight(390, 620);    //31.00 * 20
            xls.SetRowHeight(391, 620);    //31.00 * 20
            xls.SetRowHeight(392, 350);    //17.50 * 20
            xls.SetRowHeight(393, 700);    //35.00 * 20
            xls.SetRowHeight(394, 700);    //35.00 * 20
            xls.SetRowHeight(395, 700);    //35.00 * 20
            xls.SetRowHeight(396, 620);    //31.00 * 20
            xls.SetRowHeight(397, 400);    //20.00 * 20
            xls.SetRowHeight(398, 620);    //31.00 * 20
            xls.SetRowHeight(399, 620);    //31.00 * 20
            xls.SetRowHeight(400, 620);    //31.00 * 20
            xls.SetRowHeight(401, 400);    //20.00 * 20
            xls.SetRowHeight(402, 620);    //31.00 * 20
            xls.SetRowHeight(403, 630);    //31.50 * 20
            xls.SetRowHeight(404, 710);    //35.50 * 20
            xls.SetRowHeight(405, 400);    //20.00 * 20
            xls.SetRowHeight(406, 620);    //31.00 * 20
            xls.SetRowHeight(407, 620);    //31.00 * 20
            xls.SetRowHeight(408, 620);    //31.00 * 20
            xls.SetRowHeight(409, 400);    //20.00 * 20
            xls.SetRowHeight(410, 1050);    //52.50 * 20
            xls.SetRowHeight(411, 1400);    //70.00 * 20
            xls.SetRowHeight(412, 700);    //35.00 * 20
            xls.SetRowHeight(413, 700);    //35.00 * 20
            xls.SetRowHeight(414, 630);    //31.50 * 20
            xls.SetRowHeight(415, 630);    //31.50 * 20
            xls.SetRowHeight(416, 640);    //32.00 * 20

            //Merged Cells
            xls.MergeCells(62, 3, 88, 3);
            xls.MergeCells(6, 3, 55, 3);
            xls.MergeCells(143, 3, 416, 3);
            xls.MergeCells(157, 8, 176, 8);
            xls.MergeCells(182, 8, 185, 8);
            xls.MergeCells(116, 4, 142, 4);
            xls.MergeCells(89, 4, 115, 4);
            xls.MergeCells(89, 3, 115, 3);
            xls.MergeCells(116, 3, 142, 3);
            xls.MergeCells(42, 4, 55, 4);
            xls.MergeCells(56, 3, 61, 3);
            xls.MergeCells(56, 4, 61, 4);
            xls.MergeCells(6, 4, 34, 4);
            xls.MergeCells(35, 4, 40, 4);
            xls.MergeCells(196, 8, 208, 8);
            xls.MergeCells(209, 8, 217, 8);
            xls.MergeCells(62, 4, 88, 4);

            //Set the cell values
            TFlxFormat fmt;
            fmt = xls.GetCellVisibleFormatDef(1, 3);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Style = TFlxFontStyles.Bold | TFlxFontStyles.Italic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.Borders.Top.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Top.Color = TExcelColor.Automatic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(1, 3, xls.AddFormat(fmt));
            xls.SetCellValue(1, 3, "Celdas amarillas se pueden borrar, pero són útiles para el equipo de programación"
            + " y diseño que sepan dónde separar las secciones y qué subtítulos usar");

            fmt = xls.GetCellVisibleFormatDef(2, 3);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold | TFlxFontStyles.Italic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(2, 3, xls.AddFormat(fmt));
            xls.SetCellValue(2, 3, "Letras en azul dictan formato de input");

            fmt = xls.GetCellVisibleFormatDef(3, 3);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent4, 0.399975585192419);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(3, 3, xls.AddFormat(fmt));
            xls.SetCellValue(3, 3, "Cells with light purple fill are input por operations");

            fmt = xls.GetCellVisibleFormatDef(4, 3);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent4, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(4, 3, xls.AddFormat(fmt));
            xls.SetCellValue(4, 3, "Cell with dark purple fill are cells with operations");

            fmt = xls.GetCellVisibleFormatDef(5, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(5, 6, xls.AddFormat(fmt));
            xls.SetCellValue(5, 6, "INPUTS advanced");

            fmt = xls.GetCellVisibleFormatDef(6, 3);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Top.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Top.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(6, 3, xls.AddFormat(fmt));
            xls.SetCellValue(6, 3, "Años de establecimiento");

            fmt = xls.GetCellVisibleFormatDef(6, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.Borders.Top.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Top.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(6, 4, xls.AddFormat(fmt));
            xls.SetCellValue(6, 4, "Año 0");

            fmt = xls.GetCellVisibleFormatDef(6, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Top.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Top.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(6, 5, xls.AddFormat(fmt));
            xls.SetCellValue(6, 5, "Mano de obra para el germinador");

            fmt = xls.GetCellVisibleFormatDef(6, 6);
            fmt.Borders.Top.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Top.Color = TExcelColor.Automatic;
            xls.SetCellFormat(6, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(6, 7);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.Borders.Top.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Top.Color = TExcelColor.Automatic;
            xls.SetCellFormat(6, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(7, 3);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(7, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(7, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(7, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(7, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(7, 5, xls.AddFormat(fmt));
            xls.SetCellValue(7, 5, "Recolección de semillas");

            fmt = xls.GetCellVisibleFormatDef(7, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "0.00";
            xls.SetCellFormat(7, 6, xls.AddFormat(fmt));
            xls.SetCellValue(7, 6, 1.71666666666667);

            fmt = xls.GetCellVisibleFormatDef(7, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(7, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(8, 3);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(8, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(8, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(8, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(8, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(8, 5, xls.AddFormat(fmt));
            xls.SetCellValue(8, 5, "Selección de semillas");

            fmt = xls.GetCellVisibleFormatDef(8, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "0.00";
            xls.SetCellFormat(8, 6, xls.AddFormat(fmt));
            xls.SetCellValue(8, 6, 1.52243333333333);

            fmt = xls.GetCellVisibleFormatDef(8, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(8, 7, xls.AddFormat(fmt));
            xls.SetCellValue(8, 7, "INPUTS\n(Days * # of times * # of persons)");

            fmt = xls.GetCellVisibleFormatDef(9, 3);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(9, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(9, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(9, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(9, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(9, 5, xls.AddFormat(fmt));
            xls.SetCellValue(9, 5, "Construcción del semillero");

            fmt = xls.GetCellVisibleFormatDef(9, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "0.00";
            xls.SetCellFormat(9, 6, xls.AddFormat(fmt));
            xls.SetCellValue(9, 6, 4.02777777777778);

            fmt = xls.GetCellVisibleFormatDef(9, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(9, 7, xls.AddFormat(fmt));
            xls.SetCellValue(9, 7, "INPUTS\n(Days * # of times * # of persons)");

            fmt = xls.GetCellVisibleFormatDef(10, 3);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(10, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(10, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(10, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(10, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(10, 5, xls.AddFormat(fmt));
            xls.SetCellValue(10, 5, "Sostenimiento semillero - Riego");

            fmt = xls.GetCellVisibleFormatDef(10, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "0.00";
            xls.SetCellFormat(10, 6, xls.AddFormat(fmt));
            xls.SetCellValue(10, 6, 8.82);

            fmt = xls.GetCellVisibleFormatDef(10, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(10, 7, xls.AddFormat(fmt));
            xls.SetCellValue(10, 7, "INPUTS\n(Days * # of times * # of persons)");

            fmt = xls.GetCellVisibleFormatDef(11, 3);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(11, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(11, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(11, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(11, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(11, 5, xls.AddFormat(fmt));
            xls.SetCellValue(11, 5, "Otros: ");

            fmt = xls.GetCellVisibleFormatDef(11, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "0.00";
            xls.SetCellFormat(11, 6, xls.AddFormat(fmt));
            xls.SetCellValue(11, 6, 0.883333333333333);

            fmt = xls.GetCellVisibleFormatDef(11, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(11, 7, xls.AddFormat(fmt));
            xls.SetCellValue(11, 7, "INPUTS\n(Days * # of times * # of persons)");

            fmt = xls.GetCellVisibleFormatDef(12, 3);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(12, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(12, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(12, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(12, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(12, 5, xls.AddFormat(fmt));
            xls.SetCellValue(12, 5, "Mano de obra para el vivero");

            fmt = xls.GetCellVisibleFormatDef(12, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(12, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(13, 3);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(13, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(13, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(13, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(13, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(13, 5, xls.AddFormat(fmt));
            xls.SetCellValue(13, 5, "Construcción del vivero");

            fmt = xls.GetCellVisibleFormatDef(13, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "0.00";
            xls.SetCellFormat(13, 6, xls.AddFormat(fmt));
            xls.SetCellValue(13, 6, 9.61224489795918);

            fmt = xls.GetCellVisibleFormatDef(13, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(13, 7, xls.AddFormat(fmt));
            xls.SetCellValue(13, 7, "INPUTS\n(Days * # of times * # of persons)");

            fmt = xls.GetCellVisibleFormatDef(14, 3);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(14, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(14, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(14, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(14, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(14, 5, xls.AddFormat(fmt));
            xls.SetCellValue(14, 5, "Jalada y arrancada de la tierra para el vivero");

            fmt = xls.GetCellVisibleFormatDef(14, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "0.00";
            xls.SetCellFormat(14, 6, xls.AddFormat(fmt));
            xls.SetCellValue(14, 6, 8.92);

            fmt = xls.GetCellVisibleFormatDef(14, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(14, 7, xls.AddFormat(fmt));
            xls.SetCellValue(14, 7, "INPUTS\n(Days * # of times * # of persons)");

            fmt = xls.GetCellVisibleFormatDef(15, 3);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(15, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(15, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(15, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(15, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(15, 5, xls.AddFormat(fmt));
            xls.SetCellValue(15, 5, "Limpia del vivero");

            fmt = xls.GetCellVisibleFormatDef(15, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "0.00";
            xls.SetCellFormat(15, 6, xls.AddFormat(fmt));
            xls.SetCellValue(15, 6, 16.9833333333333);

            fmt = xls.GetCellVisibleFormatDef(15, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(15, 7, xls.AddFormat(fmt));
            xls.SetCellValue(15, 7, "INPUTS\n(Days * # of times * # of persons)");

            fmt = xls.GetCellVisibleFormatDef(16, 3);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(16, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(16, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(16, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(16, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(16, 5, xls.AddFormat(fmt));
            xls.SetCellValue(16, 5, "Preparacion de tierra con abono organico para llenado");

            fmt = xls.GetCellVisibleFormatDef(16, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "0.00";
            xls.SetCellFormat(16, 6, xls.AddFormat(fmt));
            xls.SetCellValue(16, 6, 6.3366);

            fmt = xls.GetCellVisibleFormatDef(16, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(16, 7, xls.AddFormat(fmt));
            xls.SetCellValue(16, 7, "INPUTS\n(Days * # of times * # of persons)");

            fmt = xls.GetCellVisibleFormatDef(17, 3);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(17, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(17, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(17, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(17, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(17, 5, xls.AddFormat(fmt));
            xls.SetCellValue(17, 5, "Llenada y encerrada de bolsas");

            fmt = xls.GetCellVisibleFormatDef(17, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "0.00";
            xls.SetCellFormat(17, 6, xls.AddFormat(fmt));
            xls.SetCellValue(17, 6, 14.78);

            fmt = xls.GetCellVisibleFormatDef(17, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(17, 7, xls.AddFormat(fmt));
            xls.SetCellValue(17, 7, "INPUTS\n(Days * # of times * # of persons)");

            fmt = xls.GetCellVisibleFormatDef(18, 3);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(18, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(18, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(18, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(18, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(18, 5, xls.AddFormat(fmt));
            xls.SetCellValue(18, 5, "Siembra de maripositas");

            fmt = xls.GetCellVisibleFormatDef(18, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "0.00";
            xls.SetCellFormat(18, 6, xls.AddFormat(fmt));
            xls.SetCellValue(18, 6, 5.45);

            fmt = xls.GetCellVisibleFormatDef(18, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(18, 7, xls.AddFormat(fmt));
            xls.SetCellValue(18, 7, "INPUTS\n(Days * # of times * # of persons)");

            fmt = xls.GetCellVisibleFormatDef(19, 3);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(19, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(19, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(19, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(19, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(19, 5, xls.AddFormat(fmt));
            xls.SetCellValue(19, 5, "Riego");

            fmt = xls.GetCellVisibleFormatDef(19, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "0.00";
            xls.SetCellFormat(19, 6, xls.AddFormat(fmt));
            xls.SetCellValue(19, 6, 24.5273333333333);

            fmt = xls.GetCellVisibleFormatDef(19, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(19, 7, xls.AddFormat(fmt));
            xls.SetCellValue(19, 7, "INPUTS\n(Days * # of times * # of persons)");

            fmt = xls.GetCellVisibleFormatDef(20, 3);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(20, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(20, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(20, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(20, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(20, 5, xls.AddFormat(fmt));
            xls.SetCellValue(20, 5, "Aplicación de foliares");

            fmt = xls.GetCellVisibleFormatDef(20, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "0.00";
            xls.SetCellFormat(20, 6, xls.AddFormat(fmt));
            xls.SetCellValue(20, 6, 2.41153333333333);

            fmt = xls.GetCellVisibleFormatDef(20, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(20, 7, xls.AddFormat(fmt));
            xls.SetCellValue(20, 7, "INPUTS\n(Days * # of times * # of persons)");

            fmt = xls.GetCellVisibleFormatDef(21, 3);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(21, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(21, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(21, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(21, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(21, 5, xls.AddFormat(fmt));
            xls.SetCellValue(21, 5, "Resiembras");

            fmt = xls.GetCellVisibleFormatDef(21, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "0.00";
            xls.SetCellFormat(21, 6, xls.AddFormat(fmt));
            xls.SetCellValue(21, 6, 1.44444444444444);

            fmt = xls.GetCellVisibleFormatDef(21, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(21, 7, xls.AddFormat(fmt));
            xls.SetCellValue(21, 7, "INPUTS\n(Days * # of times * # of persons)");

            fmt = xls.GetCellVisibleFormatDef(22, 3);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(22, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(22, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(22, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(22, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(22, 5, xls.AddFormat(fmt));
            xls.SetCellValue(22, 5, "Otros:");

            fmt = xls.GetCellVisibleFormatDef(22, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "0.00";
            xls.SetCellFormat(22, 6, xls.AddFormat(fmt));
            xls.SetCellValue(22, 6, 0.3);

            fmt = xls.GetCellVisibleFormatDef(22, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(22, 7, xls.AddFormat(fmt));
            xls.SetCellValue(22, 7, "INPUTS\n(Days * # of times * # of persons)");

            fmt = xls.GetCellVisibleFormatDef(23, 3);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(23, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(23, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(23, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(23, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.WrapText = true;
            xls.SetCellFormat(23, 5, xls.AddFormat(fmt));
            xls.SetCellValue(23, 5, "Mano de obra para preparacion del terreno y siembra");

            fmt = xls.GetCellVisibleFormatDef(23, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(23, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(24, 3);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(24, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(24, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(24, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(24, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(24, 5, xls.AddFormat(fmt));
            xls.SetCellValue(24, 5, "Limpia del terreno");

            fmt = xls.GetCellVisibleFormatDef(24, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "0.00";
            xls.SetCellFormat(24, 6, xls.AddFormat(fmt));
            xls.SetCellValue(24, 6, 18.78);

            fmt = xls.GetCellVisibleFormatDef(24, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(24, 7, xls.AddFormat(fmt));
            xls.SetCellValue(24, 7, "INPUTS\n(Days * # of times * # of persons)");

            fmt = xls.GetCellVisibleFormatDef(25, 3);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(25, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(25, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(25, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(25, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(25, 5, xls.AddFormat(fmt));
            xls.SetCellValue(25, 5, "Corte de arboles de café viejos u otros maderables");

            fmt = xls.GetCellVisibleFormatDef(25, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "0.00";
            xls.SetCellFormat(25, 6, xls.AddFormat(fmt));
            xls.SetCellValue(25, 6, 13.48);

            fmt = xls.GetCellVisibleFormatDef(25, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(25, 7, xls.AddFormat(fmt));
            xls.SetCellValue(25, 7, "INPUTS\n(Days * # of times * # of persons)");

            fmt = xls.GetCellVisibleFormatDef(26, 3);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(26, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(26, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(26, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(26, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(26, 5, xls.AddFormat(fmt));
            xls.SetCellValue(26, 5, "Recolección y acopio de madera de café");

            fmt = xls.GetCellVisibleFormatDef(26, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "0.00";
            xls.SetCellFormat(26, 6, xls.AddFormat(fmt));
            xls.SetCellValue(26, 6, 3.5);

            fmt = xls.GetCellVisibleFormatDef(26, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(26, 7, xls.AddFormat(fmt));
            xls.SetCellValue(26, 7, "INPUTS\n(Days * # of times * # of persons)");

            fmt = xls.GetCellVisibleFormatDef(27, 3);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(27, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(27, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(27, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(27, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(27, 5, xls.AddFormat(fmt));
            xls.SetCellValue(27, 5, "Pique de la madera y/o elaboración de estacas");

            fmt = xls.GetCellVisibleFormatDef(27, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "0.00";
            xls.SetCellFormat(27, 6, xls.AddFormat(fmt));
            xls.SetCellValue(27, 6, 6.12);

            fmt = xls.GetCellVisibleFormatDef(27, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(27, 7, xls.AddFormat(fmt));
            xls.SetCellValue(27, 7, "INPUTS\n(Days * # of times * # of persons)");

            fmt = xls.GetCellVisibleFormatDef(28, 3);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(28, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(28, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(28, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(28, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(28, 5, xls.AddFormat(fmt));
            xls.SetCellValue(28, 5, "Trazado Café");

            fmt = xls.GetCellVisibleFormatDef(28, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "0.00";
            xls.SetCellFormat(28, 6, xls.AddFormat(fmt));
            xls.SetCellValue(28, 6, 10.78);

            fmt = xls.GetCellVisibleFormatDef(28, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(28, 7, xls.AddFormat(fmt));
            xls.SetCellValue(28, 7, "INPUTS\n(Days * # of times * # of persons)");

            fmt = xls.GetCellVisibleFormatDef(29, 3);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(29, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(29, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(29, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(29, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(29, 5, xls.AddFormat(fmt));
            xls.SetCellValue(29, 5, "Ahoyado para la siembra");

            fmt = xls.GetCellVisibleFormatDef(29, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "0.00";
            xls.SetCellFormat(29, 6, xls.AddFormat(fmt));
            xls.SetCellValue(29, 6, 27.38);

            fmt = xls.GetCellVisibleFormatDef(29, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(29, 7, xls.AddFormat(fmt));
            xls.SetCellValue(29, 7, "INPUTS\n(Days * # of times * # of persons)");

            fmt = xls.GetCellVisibleFormatDef(30, 3);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(30, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(30, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(30, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(30, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(30, 5, xls.AddFormat(fmt));
            xls.SetCellValue(30, 5, "Llevada de las plantas del vivero (en la finca) al terreno ");

            fmt = xls.GetCellVisibleFormatDef(30, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "0.00";
            xls.SetCellFormat(30, 6, xls.AddFormat(fmt));
            xls.SetCellValue(30, 6, 12.9013333333333);

            fmt = xls.GetCellVisibleFormatDef(30, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(30, 7, xls.AddFormat(fmt));
            xls.SetCellValue(30, 7, "INPUTS\n(Days * # of times * # of persons)");

            fmt = xls.GetCellVisibleFormatDef(31, 3);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(31, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(31, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(31, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(31, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(31, 5, xls.AddFormat(fmt));
            xls.SetCellValue(31, 5, "Siembra de plantones (o plantulas)");

            fmt = xls.GetCellVisibleFormatDef(31, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "0.00";
            xls.SetCellFormat(31, 6, xls.AddFormat(fmt));
            xls.SetCellValue(31, 6, 23.34);

            fmt = xls.GetCellVisibleFormatDef(31, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(31, 7, xls.AddFormat(fmt));
            xls.SetCellValue(31, 7, "INPUTS\n(Days * # of times * # of persons)");

            fmt = xls.GetCellVisibleFormatDef(32, 3);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(32, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(32, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(32, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(32, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(32, 5, xls.AddFormat(fmt));
            xls.SetCellValue(32, 5, "Adecuación de los arboles de sombrio");

            fmt = xls.GetCellVisibleFormatDef(32, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "0.00";
            xls.SetCellFormat(32, 6, xls.AddFormat(fmt));
            xls.SetCellValue(32, 6, 13.32);

            fmt = xls.GetCellVisibleFormatDef(32, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(32, 7, xls.AddFormat(fmt));
            xls.SetCellValue(32, 7, "INPUTS\n(Days * # of times * # of persons)");

            fmt = xls.GetCellVisibleFormatDef(33, 3);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(33, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(33, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(33, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(33, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(33, 5, xls.AddFormat(fmt));
            xls.SetCellValue(33, 5, "Preparación de abonos orgánicos");

            fmt = xls.GetCellVisibleFormatDef(33, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "0.00";
            xls.SetCellFormat(33, 6, xls.AddFormat(fmt));
            xls.SetCellValue(33, 6, 4.66);

            fmt = xls.GetCellVisibleFormatDef(33, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(33, 7, xls.AddFormat(fmt));
            xls.SetCellValue(33, 7, "INPUTS\n(Days * # of times * # of persons)");

            fmt = xls.GetCellVisibleFormatDef(34, 3);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(34, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(34, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(34, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(34, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(34, 5, xls.AddFormat(fmt));
            xls.SetCellValue(34, 5, "Otros");

            fmt = xls.GetCellVisibleFormatDef(34, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.Format = "0.00";
            xls.SetCellFormat(34, 6, xls.AddFormat(fmt));
            xls.SetCellValue(34, 6, 1.2998);

            fmt = xls.GetCellVisibleFormatDef(34, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(34, 7, xls.AddFormat(fmt));
            xls.SetCellValue(34, 7, "INPUTS\n(Days * # of times * # of persons)");

            fmt = xls.GetCellVisibleFormatDef(35, 3);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(35, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(35, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.Borders.Top.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Top.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(35, 4, xls.AddFormat(fmt));
            xls.SetCellValue(35, 4, "Año 1");

            fmt = xls.GetCellVisibleFormatDef(35, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Top.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Top.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.WrapText = true;
            xls.SetCellFormat(35, 5, xls.AddFormat(fmt));
            xls.SetCellValue(35, 5, "Mano de obra para la plantilla o levante");

            fmt = xls.GetCellVisibleFormatDef(35, 6);
            fmt.Borders.Top.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Top.Color = TExcelColor.Automatic;
            xls.SetCellFormat(35, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(35, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.Borders.Top.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Top.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(35, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(36, 3);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(36, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(36, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(36, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(36, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(36, 5, xls.AddFormat(fmt));
            xls.SetCellValue(36, 5, "Desyerbe periodico ");

            fmt = xls.GetCellVisibleFormatDef(36, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "0.00";
            xls.SetCellFormat(36, 6, xls.AddFormat(fmt));
            xls.SetCellValue(36, 6, 47.4285714285714);

            fmt = xls.GetCellVisibleFormatDef(36, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(36, 7, xls.AddFormat(fmt));
            xls.SetCellValue(36, 7, "INPUTS\n(Days * # of times * # of persons)");

            fmt = xls.GetCellVisibleFormatDef(37, 3);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(37, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(37, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(37, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(37, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(37, 5, xls.AddFormat(fmt));
            xls.SetCellValue(37, 5, "Aplicación de abonos orgánicos para levante");

            fmt = xls.GetCellVisibleFormatDef(37, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "0.00";
            xls.SetCellFormat(37, 6, xls.AddFormat(fmt));
            xls.SetCellValue(37, 6, 5.32);

            fmt = xls.GetCellVisibleFormatDef(37, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(37, 7, xls.AddFormat(fmt));
            xls.SetCellValue(37, 7, "INPUTS\n(Days * # of times * # of persons)");

            fmt = xls.GetCellVisibleFormatDef(38, 3);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(38, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(38, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(38, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(38, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(38, 5, xls.AddFormat(fmt));
            xls.SetCellValue(38, 5, "Aplicación de abonos químicos para levante");

            fmt = xls.GetCellVisibleFormatDef(38, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "0.00";
            xls.SetCellFormat(38, 6, xls.AddFormat(fmt));
            xls.SetCellValue(38, 6, 0.24);

            fmt = xls.GetCellVisibleFormatDef(38, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(38, 7, xls.AddFormat(fmt));
            xls.SetCellValue(38, 7, "INPUTS\n(Days * # of times * # of persons)");

            fmt = xls.GetCellVisibleFormatDef(39, 3);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(39, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(39, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(39, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(39, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(39, 5, xls.AddFormat(fmt));
            xls.SetCellValue(39, 5, "Aplicación de foliares para fertilización y control roya");

            fmt = xls.GetCellVisibleFormatDef(39, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "0.00";
            xls.SetCellFormat(39, 6, xls.AddFormat(fmt));
            xls.SetCellValue(39, 6, 6.2);

            fmt = xls.GetCellVisibleFormatDef(39, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(39, 7, xls.AddFormat(fmt));
            xls.SetCellValue(39, 7, "INPUTS\n(Days * # of times * # of persons)");

            fmt = xls.GetCellVisibleFormatDef(40, 3);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(40, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(40, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(40, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(40, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(40, 5, xls.AddFormat(fmt));
            xls.SetCellValue(40, 5, "Otros:");

            fmt = xls.GetCellVisibleFormatDef(40, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.Format = "0.00";
            xls.SetCellFormat(40, 6, xls.AddFormat(fmt));
            xls.SetCellValue(40, 6, 2.1);

            fmt = xls.GetCellVisibleFormatDef(40, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(40, 7, xls.AddFormat(fmt));
            xls.SetCellValue(40, 7, "INPUTS\n(Days * # of times * # of persons)");

            fmt = xls.GetCellVisibleFormatDef(41, 3);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(41, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(41, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.Borders.Top.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Top.Color = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(41, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(41, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 320;
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Style = TFlxFontStyles.Bold | TFlxFontStyles.Italic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Top.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Top.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(41, 5, xls.AddFormat(fmt));
            xls.SetCellValue(41, 5, "Transporte");

            fmt = xls.GetCellVisibleFormatDef(41, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Top.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Top.Color = TExcelColor.Automatic;
            fmt.Format = "0.00";
            xls.SetCellFormat(41, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(41, 7);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.Borders.Top.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Top.Color = TExcelColor.Automatic;
            xls.SetCellFormat(41, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(42, 3);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(42, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(42, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(42, 4, xls.AddFormat(fmt));
            xls.SetCellValue(42, 4, "Año 0");

            fmt = xls.GetCellVisibleFormatDef(42, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Style = TFlxFontStyles.Bold | TFlxFontStyles.Italic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(42, 5, xls.AddFormat(fmt));
            xls.SetCellValue(42, 5, "Transporte relacionado con actividades o insumos para el \nGERMINADOR");

            fmt = xls.GetCellVisibleFormatDef(42, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "0.00";
            xls.SetCellFormat(42, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(42, 7);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            xls.SetCellFormat(42, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(43, 3);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(43, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(43, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(43, 4, xls.AddFormat(fmt));
            xls.SetCellValue(43, 5, "ir a comprar la semilla");

            fmt = xls.GetCellVisibleFormatDef(43, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "0.00";
            xls.SetCellFormat(43, 6, xls.AddFormat(fmt));
            xls.SetCellValue(43, 6, 79.78);

            fmt = xls.GetCellVisibleFormatDef(43, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(43, 7, xls.AddFormat(fmt));
            xls.SetCellValue(43, 7, "INPUTS\n#");

            fmt = xls.GetCellVisibleFormatDef(44, 3);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(44, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(44, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(44, 4, xls.AddFormat(fmt));
            xls.SetCellValue(44, 5, "Llevada madera");

            fmt = xls.GetCellVisibleFormatDef(44, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "0.00";
            xls.SetCellFormat(44, 6, xls.AddFormat(fmt));
            xls.SetCellValue(44, 6, 113.84);

            fmt = xls.GetCellVisibleFormatDef(44, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(44, 7, xls.AddFormat(fmt));
            xls.SetCellValue(44, 7, "INPUTS\n#");

            fmt = xls.GetCellVisibleFormatDef(45, 3);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(45, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(45, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(45, 4, xls.AddFormat(fmt));
            xls.SetCellValue(45, 5, "Llevada arena");

            fmt = xls.GetCellVisibleFormatDef(45, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "0.00";
            xls.SetCellFormat(45, 6, xls.AddFormat(fmt));
            xls.SetCellValue(45, 6, 172.32);

            fmt = xls.GetCellVisibleFormatDef(45, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(45, 7, xls.AddFormat(fmt));
            xls.SetCellValue(45, 7, "INPUTS\n#");

            fmt = xls.GetCellVisibleFormatDef(46, 3);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(46, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(46, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(46, 4, xls.AddFormat(fmt));
            xls.SetCellValue(46, 5, "Otro(s):");

            fmt = xls.GetCellVisibleFormatDef(46, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "0.00";
            xls.SetCellFormat(46, 6, xls.AddFormat(fmt));
            xls.SetCellValue(46, 6, 0);

            fmt = xls.GetCellVisibleFormatDef(46, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(46, 7, xls.AddFormat(fmt));
            xls.SetCellValue(46, 7, "INPUTS\n#");

            fmt = xls.GetCellVisibleFormatDef(47, 3);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(47, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(47, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(47, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(47, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Style = TFlxFontStyles.Bold | TFlxFontStyles.Italic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(47, 5, xls.AddFormat(fmt));
            xls.SetCellValue(47, 5, "Transporte relacionado con actividades o insumos para el \nVIVERO o ALMÁCIGO");

            fmt = xls.GetCellVisibleFormatDef(47, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "0.00";
            xls.SetCellFormat(47, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(47, 7);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            xls.SetCellFormat(47, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(48, 3);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(48, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(48, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(48, 4, xls.AddFormat(fmt));
            xls.SetCellValue(48, 5, "Jalada de tierra");

            fmt = xls.GetCellVisibleFormatDef(48, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "0.00";
            xls.SetCellFormat(48, 6, xls.AddFormat(fmt));
            xls.SetCellValue(48, 6, 475.84);

            fmt = xls.GetCellVisibleFormatDef(48, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(48, 7, xls.AddFormat(fmt));
            xls.SetCellValue(48, 7, "INPUTS\n#");

            fmt = xls.GetCellVisibleFormatDef(49, 3);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(49, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(49, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(49, 4, xls.AddFormat(fmt));
            xls.SetCellValue(49, 5, "Ir a comprar bolsas y otros insumos para el vivero");

            fmt = xls.GetCellVisibleFormatDef(49, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "0.00";
            xls.SetCellFormat(49, 6, xls.AddFormat(fmt));
            xls.SetCellValue(49, 6, new TFormula("=F59/4"));

            fmt = xls.GetCellVisibleFormatDef(49, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(49, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(50, 3);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(50, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(50, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(50, 4, xls.AddFormat(fmt));
            xls.SetCellValue(50, 5, "Otro(s)");

            fmt = xls.GetCellVisibleFormatDef(50, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "0.00";
            xls.SetCellFormat(50, 6, xls.AddFormat(fmt));
            xls.SetCellValue(50, 6, 0);

            fmt = xls.GetCellVisibleFormatDef(50, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(50, 7, xls.AddFormat(fmt));
            xls.SetCellValue(50, 7, "INPUTS\n#");

            fmt = xls.GetCellVisibleFormatDef(51, 3);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(51, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(51, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(51, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(51, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Style = TFlxFontStyles.Bold | TFlxFontStyles.Italic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(51, 5, xls.AddFormat(fmt));
            xls.SetCellValue(51, 5, "Transporte relacionado con actividades o insumos para el \nPREPARACIÓN DEL TERRENO"
            + " Y SIEMBRA");

            fmt = xls.GetCellVisibleFormatDef(51, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "0.00";
            xls.SetCellFormat(51, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(51, 7);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            xls.SetCellFormat(51, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(52, 3);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(52, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(52, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(52, 4, xls.AddFormat(fmt));
            xls.SetCellValue(52, 5, "Llevada de leña");

            fmt = xls.GetCellVisibleFormatDef(52, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "0.00";
            xls.SetCellFormat(52, 6, xls.AddFormat(fmt));
            xls.SetCellValue(52, 6, 266.4);

            fmt = xls.GetCellVisibleFormatDef(52, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(52, 7, xls.AddFormat(fmt));
            xls.SetCellValue(52, 7, "INPUTS\n#");

            fmt = xls.GetCellVisibleFormatDef(53, 3);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(53, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(53, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(53, 4, xls.AddFormat(fmt));
            xls.SetCellValue(53, 5, "Lleva del abono");

            fmt = xls.GetCellVisibleFormatDef(53, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "0.00";
            xls.SetCellFormat(53, 6, xls.AddFormat(fmt));
            xls.SetCellValue(53, 6, 142.12);

            fmt = xls.GetCellVisibleFormatDef(53, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(53, 7, xls.AddFormat(fmt));
            xls.SetCellValue(53, 7, "INPUTS\n#");

            fmt = xls.GetCellVisibleFormatDef(54, 3);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(54, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(54, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(54, 4, xls.AddFormat(fmt));
            xls.SetCellValue(54, 5, "Llevar plantas del vivero al campo");

            fmt = xls.GetCellVisibleFormatDef(54, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "0.00";
            xls.SetCellFormat(54, 6, xls.AddFormat(fmt));
            xls.SetCellValue(54, 6, 1817.6);

            fmt = xls.GetCellVisibleFormatDef(54, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(54, 7, xls.AddFormat(fmt));
            xls.SetCellValue(54, 7, "INPUTS\n#");

            fmt = xls.GetCellVisibleFormatDef(55, 3);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(55, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(55, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(55, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(55, 5);
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            xls.SetCellFormat(55, 5, xls.AddFormat(fmt));
            xls.SetCellValue(55, 5, "Otro(s)");

            fmt = xls.GetCellVisibleFormatDef(55, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.Format = "0.00";
            xls.SetCellFormat(55, 6, xls.AddFormat(fmt));
            xls.SetCellValue(55, 6, 0);

            fmt = xls.GetCellVisibleFormatDef(55, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(55, 7, xls.AddFormat(fmt));
            xls.SetCellValue(55, 7, "INPUTS\n#");

            fmt = xls.GetCellVisibleFormatDef(56, 3);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.Borders.Top.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Top.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(56, 3, xls.AddFormat(fmt));
            xls.SetCellValue(56, 3, "Transporte Años productivos");

            fmt = xls.GetCellVisibleFormatDef(56, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.Borders.Top.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Top.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(56, 4, xls.AddFormat(fmt));
            xls.SetCellValue(56, 4, "Años 2-8");

            fmt = xls.GetCellVisibleFormatDef(56, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Style = TFlxFontStyles.Bold | TFlxFontStyles.Italic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Top.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Top.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(56, 5, xls.AddFormat(fmt));
            xls.SetCellValue(56, 5, "Otros gastos en transporte en términos anuales");

            fmt = xls.GetCellVisibleFormatDef(56, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Top.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Top.Color = TExcelColor.Automatic;
            fmt.Format = "0.00";
            xls.SetCellFormat(56, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(56, 7);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.Borders.Top.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Top.Color = TExcelColor.Automatic;
            xls.SetCellFormat(56, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(57, 3);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(57, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(57, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(57, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(57, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(57, 5, xls.AddFormat(fmt));
            xls.SetCellValue(57, 5, "Transporte equipo y herramientas");

            fmt = xls.GetCellVisibleFormatDef(57, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "0.00";
            xls.SetCellFormat(57, 6, xls.AddFormat(fmt));
            xls.SetCellValue(57, 6, 439.413333333333);

            fmt = xls.GetCellVisibleFormatDef(57, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(57, 7, xls.AddFormat(fmt));
            xls.SetCellValue(57, 7, "INPUTS\n#");

            fmt = xls.GetCellVisibleFormatDef(58, 3);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(58, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(58, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(58, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(58, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.WrapText = true;
            xls.SetCellFormat(58, 5, xls.AddFormat(fmt));
            xls.SetCellValue(58, 5, "Transporte mano de obra (no pagada en el jornal)");

            fmt = xls.GetCellVisibleFormatDef(58, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "0.00";
            xls.SetCellFormat(58, 6, xls.AddFormat(fmt));
            xls.SetCellValue(58, 6, 1195.8);

            fmt = xls.GetCellVisibleFormatDef(58, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(58, 7, xls.AddFormat(fmt));
            xls.SetCellValue(58, 7, "INPUTS\n#");

            fmt = xls.GetCellVisibleFormatDef(59, 3);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(59, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(59, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(59, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(59, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.WrapText = true;
            xls.SetCellFormat(59, 5, xls.AddFormat(fmt));
            xls.SetCellValue(59, 5, "Transporte de la cosecha al centro de acopio o asociación  ");

            fmt = xls.GetCellVisibleFormatDef(59, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent2, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0.00";
            xls.SetCellFormat(59, 6, xls.AddFormat(fmt));
            xls.SetCellValue(59, 6, new TFormula("='Inputs 1.0_metric_currency'!D17"));

            fmt = xls.GetCellVisibleFormatDef(59, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(59, 7, xls.AddFormat(fmt));
            xls.SetCellValue(59, 7, "INPUTS\n#");

            fmt = xls.GetCellVisibleFormatDef(60, 3);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(60, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(60, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(60, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(60, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.WrapText = true;
            xls.SetCellFormat(60, 5, xls.AddFormat(fmt));
            xls.SetCellValue(60, 5, "Transporte para ir a supervisar actividades (Limpias, manejos, podas, obras conservación)");

            fmt = xls.GetCellVisibleFormatDef(60, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "0.00";
            xls.SetCellFormat(60, 6, xls.AddFormat(fmt));
            xls.SetCellValue(60, 6, 751.2);

            fmt = xls.GetCellVisibleFormatDef(60, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(60, 7, xls.AddFormat(fmt));
            xls.SetCellValue(60, 7, "INPUTS\n#");

            fmt = xls.GetCellVisibleFormatDef(61, 3);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(61, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(61, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(61, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(61, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            xls.SetCellFormat(61, 5, xls.AddFormat(fmt));
            xls.SetCellValue(61, 5, "Otro(s) transportes no considerados:");

            fmt = xls.GetCellVisibleFormatDef(61, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.Format = "0.00";
            xls.SetCellFormat(61, 6, xls.AddFormat(fmt));
            xls.SetCellValue(61, 6, 0);

            fmt = xls.GetCellVisibleFormatDef(61, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(61, 7, xls.AddFormat(fmt));
            xls.SetCellValue(61, 7, "INPUTS\n#");

            fmt = xls.GetCellVisibleFormatDef(62, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.Borders.Top.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Top.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(62, 3, xls.AddFormat(fmt));
            xls.SetCellValue(62, 3, "Árbol joven");

            fmt = xls.GetCellVisibleFormatDef(62, 4);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.Borders.Top.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Top.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(62, 4, xls.AddFormat(fmt));
            xls.SetCellValue(62, 4, "Años 2 y 3");

            fmt = xls.GetCellVisibleFormatDef(62, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Top.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Top.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.WrapText = true;
            xls.SetCellFormat(62, 5, xls.AddFormat(fmt));
            xls.SetCellValue(62, 5, "Mano de obra para mantenimiento");

            fmt = xls.GetCellVisibleFormatDef(62, 6);
            fmt.Borders.Top.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Top.Color = TExcelColor.Automatic;
            xls.SetCellFormat(62, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(62, 7);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.Borders.Top.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Top.Color = TExcelColor.Automatic;
            xls.SetCellFormat(62, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(63, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(63, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(63, 4);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(63, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(63, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(63, 5, xls.AddFormat(fmt));
            xls.SetCellValue(63, 5, "Desyerbe para mantenimiento");

            fmt = xls.GetCellVisibleFormatDef(63, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "0.00";
            xls.SetCellFormat(63, 6, xls.AddFormat(fmt));
            xls.SetCellValue(63, 6, 40.34);

            fmt = xls.GetCellVisibleFormatDef(63, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(63, 7, xls.AddFormat(fmt));
            xls.SetCellValue(63, 7, "INPUTS\n(Days * # of times * # of persons)");

            fmt = xls.GetCellVisibleFormatDef(63, 8);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(63, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(64, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(64, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(64, 4);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(64, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(64, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(64, 5, xls.AddFormat(fmt));
            xls.SetCellValue(64, 5, "Desyerbe químico para mantenimiento");

            fmt = xls.GetCellVisibleFormatDef(64, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "0.00";
            xls.SetCellFormat(64, 6, xls.AddFormat(fmt));
            xls.SetCellValue(64, 6, 0.04);

            fmt = xls.GetCellVisibleFormatDef(64, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(64, 7, xls.AddFormat(fmt));
            xls.SetCellValue(64, 7, "INPUTS\n(Days * # of times * # of persons)");

            fmt = xls.GetCellVisibleFormatDef(65, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(65, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(65, 4);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(65, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(65, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(65, 5, xls.AddFormat(fmt));
            xls.SetCellValue(65, 5, "Aplicación de abonos orgánicos para mantenimiento");

            fmt = xls.GetCellVisibleFormatDef(65, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "0.00";
            xls.SetCellFormat(65, 6, xls.AddFormat(fmt));
            xls.SetCellValue(65, 6, 5.75);

            fmt = xls.GetCellVisibleFormatDef(65, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(65, 7, xls.AddFormat(fmt));
            xls.SetCellValue(65, 7, "INPUTS\n(Days * # of times * # of persons)");

            fmt = xls.GetCellVisibleFormatDef(66, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(66, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(66, 4);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(66, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(66, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(66, 5, xls.AddFormat(fmt));
            xls.SetCellValue(66, 5, "Aplicación de abonos químicos para mantenimiento");

            fmt = xls.GetCellVisibleFormatDef(66, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "0.00";
            xls.SetCellFormat(66, 6, xls.AddFormat(fmt));
            xls.SetCellValue(66, 6, 0);

            fmt = xls.GetCellVisibleFormatDef(66, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(66, 7, xls.AddFormat(fmt));
            xls.SetCellValue(66, 7, "INPUTS\n(Days * # of times * # of persons)");

            fmt = xls.GetCellVisibleFormatDef(67, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(67, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(67, 4);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(67, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(67, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(67, 5, xls.AddFormat(fmt));
            xls.SetCellValue(67, 5, "Aplicación de foliares para fertilización y control roya");

            fmt = xls.GetCellVisibleFormatDef(67, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "0.00";
            xls.SetCellFormat(67, 6, xls.AddFormat(fmt));
            xls.SetCellValue(67, 6, 3.4);

            fmt = xls.GetCellVisibleFormatDef(67, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(67, 7, xls.AddFormat(fmt));
            xls.SetCellValue(67, 7, "INPUTS\n(Days * # of times * # of persons)");

            fmt = xls.GetCellVisibleFormatDef(68, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(68, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(68, 4);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(68, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(68, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(68, 5, xls.AddFormat(fmt));
            xls.SetCellValue(68, 5, "Construcción de barreras vivas (rompe-vientos)");

            fmt = xls.GetCellVisibleFormatDef(68, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "0.00";
            xls.SetCellFormat(68, 6, xls.AddFormat(fmt));
            xls.SetCellValue(68, 6, 4);

            fmt = xls.GetCellVisibleFormatDef(68, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(68, 7, xls.AddFormat(fmt));
            xls.SetCellValue(68, 7, "INPUTS\n(Days * # of times * # of persons)");

            fmt = xls.GetCellVisibleFormatDef(69, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(69, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(69, 4);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(69, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(69, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(69, 5, xls.AddFormat(fmt));
            xls.SetCellValue(69, 5, "Podas de árboles de sombra (sostenimiento)");

            fmt = xls.GetCellVisibleFormatDef(69, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "0.00";
            xls.SetCellFormat(69, 6, xls.AddFormat(fmt));
            xls.SetCellValue(69, 6, 13);

            fmt = xls.GetCellVisibleFormatDef(69, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(69, 7, xls.AddFormat(fmt));
            xls.SetCellValue(69, 7, "INPUTS\n(Days * # of times * # of persons)");

            fmt = xls.GetCellVisibleFormatDef(70, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(70, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(70, 4);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(70, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(70, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.left;
            xls.SetCellFormat(70, 5, xls.AddFormat(fmt));
            xls.SetCellValue(70, 5, "Control de Broca (re-re, repela, fumigaciones)");

            fmt = xls.GetCellVisibleFormatDef(70, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "0.00";
            xls.SetCellFormat(70, 6, xls.AddFormat(fmt));
            xls.SetCellValue(70, 6, 0.3);

            fmt = xls.GetCellVisibleFormatDef(70, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(70, 7, xls.AddFormat(fmt));
            xls.SetCellValue(70, 7, "INPUTS\n(Days * # of times * # of persons)");

            fmt = xls.GetCellVisibleFormatDef(71, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(71, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(71, 4);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(71, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(71, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(71, 5, xls.AddFormat(fmt));
            xls.SetCellValue(71, 5, "Manejo de tejido (desrrame o podas del café)");

            fmt = xls.GetCellVisibleFormatDef(71, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "0.00";
            xls.SetCellFormat(71, 6, xls.AddFormat(fmt));
            xls.SetCellValue(71, 6, 8.9);

            fmt = xls.GetCellVisibleFormatDef(71, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(71, 7, xls.AddFormat(fmt));
            xls.SetCellValue(71, 7, "INPUTS\n(Days * # of times * # of persons)");

            fmt = xls.GetCellVisibleFormatDef(72, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(72, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(72, 4);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(72, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(72, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.left;
            xls.SetCellFormat(72, 5, xls.AddFormat(fmt));
            xls.SetCellValue(72, 5, "Otros:");

            fmt = xls.GetCellVisibleFormatDef(72, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "0.00";
            xls.SetCellFormat(72, 6, xls.AddFormat(fmt));
            xls.SetCellValue(72, 6, 7.84);

            fmt = xls.GetCellVisibleFormatDef(72, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(72, 7, xls.AddFormat(fmt));
            xls.SetCellValue(72, 7, "INPUTS\n(Days * # of times * # of persons)");

            fmt = xls.GetCellVisibleFormatDef(73, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(73, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(73, 4);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(73, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(73, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.WrapText = true;
            xls.SetCellFormat(73, 5, xls.AddFormat(fmt));
            xls.SetCellValue(73, 5, "Mano de obra para cosecha");

            fmt = xls.GetCellVisibleFormatDef(73, 7);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            xls.SetCellFormat(73, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(74, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(74, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(74, 4);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(74, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(74, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Locked = false;
            fmt.WrapText = true;
            xls.SetCellFormat(74, 5, xls.AddFormat(fmt));

            TRTFRun[] Runs;
            Runs = new TRTFRun[2];
            Runs[0].FirstChar = 8;
            TFlxFont fnt;
            fnt = xls.GetDefaultFont;
            fnt.Name = "Arial";
            fnt.Color = TUIColor.FromArgb(0x00, 0xB0, 0x50);
            fnt.Style = TFlxFontStyles.Bold;
            fnt.Scheme = TFontScheme.None;
            Runs[0].FontIndex = xls.AddFont(fnt);
            Runs[1].FirstChar = 17;
            fnt = xls.GetDefaultFont;
            fnt.Name = "Arial";
            fnt.Color = TExcelColor.Automatic;
            fnt.Scheme = TFontScheme.None;
            Runs[1].FontIndex = xls.AddFont(fnt);
            xls.SetCellValue(74, 5, new TRichString("Cuantos Quintales cosecha por hectarea", Runs, xls));
            //We could also have used: xls.SetCellFromHtml(74, 5, "Cuantos&nbsp;<font color = '#00b050'><b>Quintales</b></font>&nbsp;cosecha por hectarea")

            xls.SetCellValue(74, 6, 3.4);

            fmt = xls.GetCellVisibleFormatDef(74, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(74, 7, xls.AddFormat(fmt));
            xls.SetCellValue(74, 7, "INPUTS\n#");

            fmt = xls.GetCellVisibleFormatDef(75, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(75, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(75, 4);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(75, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(75, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(75, 5, xls.AddFormat(fmt));
            xls.SetCellValue(75, 5, "TOTAL de tiempo recogiendo café");
            xls.SetCellValue(75, 6, 25);

            fmt = xls.GetCellVisibleFormatDef(75, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(75, 7, xls.AddFormat(fmt));
            xls.SetCellValue(75, 7, "INPUTS\n(Days * # of times * # of persons)");

            fmt = xls.GetCellVisibleFormatDef(76, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(76, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(76, 4);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(76, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(76, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(76, 5, xls.AddFormat(fmt));
            xls.SetCellValue(76, 5, "Otras actividades relacionadas con la cosecha");
            xls.SetCellValue(76, 6, 0);

            fmt = xls.GetCellVisibleFormatDef(76, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(76, 7, xls.AddFormat(fmt));
            xls.SetCellValue(76, 7, "INPUTS\n#");

            fmt = xls.GetCellVisibleFormatDef(77, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(77, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(77, 4);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(77, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(77, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(77, 5, xls.AddFormat(fmt));
            xls.SetCellValue(77, 5, "¿Cuántos KILOS de CEREZA recoge POR HECTÁREA?");

            fmt = xls.GetCellVisibleFormatDef(77, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "0.00";
            xls.SetCellFormat(77, 6, xls.AddFormat(fmt));
            xls.SetCellValue(77, 6, 3.4);

            fmt = xls.GetCellVisibleFormatDef(77, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(77, 7, xls.AddFormat(fmt));
            xls.SetCellValue(77, 7, "INPUTS\n#");

            fmt = xls.GetCellVisibleFormatDef(78, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(78, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(78, 4);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(78, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(78, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.left;
            xls.SetCellFormat(78, 5, xls.AddFormat(fmt));

            Runs = new TRTFRun[2];
            Runs[0].FirstChar = 9;
            fnt = xls.GetDefaultFont;
            fnt.Name = "Arial";
            fnt.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fnt.Style = TFlxFontStyles.Bold;
            fnt.Scheme = TFontScheme.None;
            Runs[0].FontIndex = xls.AddFont(fnt);
            Runs[1].FirstChar = 18;
            fnt = xls.GetDefaultFont;
            fnt.Name = "Arial";
            fnt.Color = TExcelColor.Automatic;
            fnt.Scheme = TFontScheme.None;
            Runs[1].FontIndex = xls.AddFont(fnt);
            xls.SetCellValue(78, 5, new TRichString("¿Cuántas QUINTALES de PERGAMINO SECO recoge POR HECTÁREA?", Runs, xls));
            //We could also have used: xls.SetCellFromHtml(78, 5, "&iquest;Cu&aacute;ntas&nbsp;<font color = 'red'><b>QUINTALES</b></font>&nbsp;de PERGAMINO"
            //+" SECO recoge POR HECT&Aacute;REA?")


    fmt = xls.GetCellVisibleFormatDef(78, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "0.00";
            xls.SetCellFormat(78, 6, xls.AddFormat(fmt));
            xls.SetCellValue(78, 6, new TFormula("=Proportions!J5"));

            fmt = xls.GetCellVisibleFormatDef(78, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(78, 7, xls.AddFormat(fmt));
            xls.SetCellValue(78, 7, "INPUTS\n#");

            fmt = xls.GetCellVisibleFormatDef(79, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(79, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(79, 4);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(79, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(79, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.WrapText = true;
            xls.SetCellFormat(79, 5, xls.AddFormat(fmt));

            Runs = new TRTFRun[1];
            Runs[0].FirstChar = 31;
            fnt = xls.GetDefaultFont;
            fnt.Name = "Arial";
            fnt.Size20 = 280;
            fnt.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fnt.Style = TFlxFontStyles.Bold;
            fnt.Scheme = TFontScheme.None;
            Runs[0].FontIndex = xls.AddFont(fnt);
            xls.SetCellValue(79, 5, new TRichString("Mano de obra para el beneficio (en Horas)", Runs, xls));
            //We could also have used: xls.SetCellFromHtml(79, 5, "Mano de obra para el beneficio&nbsp;<font color = 'red'>(en Horas)</font>")


            fmt = xls.GetCellVisibleFormatDef(79, 7);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            xls.SetCellFormat(79, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(80, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(80, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(80, 4);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(80, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(80, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(80, 5, xls.AddFormat(fmt));
            xls.SetCellValue(80, 5, "Despulpado y Fermentado");

            fmt = xls.GetCellVisibleFormatDef(80, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "0.00";
            xls.SetCellFormat(80, 6, xls.AddFormat(fmt));
            xls.SetCellValue(80, 6, 3);

            fmt = xls.GetCellVisibleFormatDef(80, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(80, 7, xls.AddFormat(fmt));
            xls.SetCellValue(80, 7, "INPUTS\n(Hours* # of times * # of persons)");

            fmt = xls.GetCellVisibleFormatDef(81, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(81, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(81, 4);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(81, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(81, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(81, 5, xls.AddFormat(fmt));
            xls.SetCellValue(81, 5, "Lavado (incluye rebalse)");

            fmt = xls.GetCellVisibleFormatDef(81, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "0.00";
            xls.SetCellFormat(81, 6, xls.AddFormat(fmt));
            xls.SetCellValue(81, 6, 3);

            fmt = xls.GetCellVisibleFormatDef(81, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(81, 7, xls.AddFormat(fmt));
            xls.SetCellValue(81, 7, "INPUTS\n(Hours* # of times * # of persons)");

            fmt = xls.GetCellVisibleFormatDef(82, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(82, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(82, 4);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(82, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(82, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(82, 5, xls.AddFormat(fmt));
            xls.SetCellValue(82, 5, "Secado");

            fmt = xls.GetCellVisibleFormatDef(82, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "0.00";
            xls.SetCellFormat(82, 6, xls.AddFormat(fmt));
            xls.SetCellValue(82, 6, 5.8);

            fmt = xls.GetCellVisibleFormatDef(82, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(82, 7, xls.AddFormat(fmt));
            xls.SetCellValue(82, 7, "INPUTS\n(Hours* # of times * # of persons)");

            fmt = xls.GetCellVisibleFormatDef(83, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(83, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(83, 4);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(83, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(83, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(83, 5, xls.AddFormat(fmt));
            xls.SetCellValue(83, 5, "Zarandeo");

            fmt = xls.GetCellVisibleFormatDef(83, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "0.00";
            xls.SetCellFormat(83, 6, xls.AddFormat(fmt));
            xls.SetCellValue(83, 6, 1.2);

            fmt = xls.GetCellVisibleFormatDef(83, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(83, 7, xls.AddFormat(fmt));
            xls.SetCellValue(83, 7, "INPUTS\n(Hours* # of times * # of persons)");

            fmt = xls.GetCellVisibleFormatDef(84, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(84, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(84, 4);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(84, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(84, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.left;
            xls.SetCellFormat(84, 5, xls.AddFormat(fmt));
            xls.SetCellValue(84, 5, "Escojo/selección");

            fmt = xls.GetCellVisibleFormatDef(84, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "0.00";
            xls.SetCellFormat(84, 6, xls.AddFormat(fmt));
            xls.SetCellValue(84, 6, 1.8);

            fmt = xls.GetCellVisibleFormatDef(84, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(84, 7, xls.AddFormat(fmt));
            xls.SetCellValue(84, 7, "INPUTS\n(Hours* # of times * # of persons)");

            fmt = xls.GetCellVisibleFormatDef(85, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(85, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(85, 4);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(85, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(85, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(85, 5, xls.AddFormat(fmt));
            xls.SetCellValue(85, 5, "Almacenamiento");
            xls.SetCellValue(85, 6, 1);

            fmt = xls.GetCellVisibleFormatDef(85, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(85, 7, xls.AddFormat(fmt));
            xls.SetCellValue(85, 7, "INPUTS\n(Hours* # of times * # of persons)");

            fmt = xls.GetCellVisibleFormatDef(86, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(86, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(86, 4);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(86, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(86, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(86, 5, xls.AddFormat(fmt));
            xls.SetCellValue(86, 5, "Aguas Miel");

            fmt = xls.GetCellVisibleFormatDef(86, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "0.00";
            xls.SetCellFormat(86, 6, xls.AddFormat(fmt));
            xls.SetCellValue(86, 6, 0.28);

            fmt = xls.GetCellVisibleFormatDef(86, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(86, 7, xls.AddFormat(fmt));
            xls.SetCellValue(86, 7, "INPUTS\n(Hours* # of times * # of persons)");

            fmt = xls.GetCellVisibleFormatDef(87, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(87, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(87, 4);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(87, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(87, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(87, 5, xls.AddFormat(fmt));
            xls.SetCellValue(87, 5, "Manejo de pulpa");

            fmt = xls.GetCellVisibleFormatDef(87, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "0.00";
            xls.SetCellFormat(87, 6, xls.AddFormat(fmt));
            xls.SetCellValue(87, 6, 1.9);

            fmt = xls.GetCellVisibleFormatDef(87, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(87, 7, xls.AddFormat(fmt));
            xls.SetCellValue(87, 7, "INPUTS\n(Hours* # of times * # of persons)");

            fmt = xls.GetCellVisibleFormatDef(88, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(88, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(88, 4);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(88, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(88, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            xls.SetCellFormat(88, 5, xls.AddFormat(fmt));
            xls.SetCellValue(88, 5, "Otros");

            fmt = xls.GetCellVisibleFormatDef(88, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.Format = "0.00";
            xls.SetCellFormat(88, 6, xls.AddFormat(fmt));
            xls.SetCellValue(88, 6, 0.1);

            fmt = xls.GetCellVisibleFormatDef(88, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(88, 7, xls.AddFormat(fmt));
            xls.SetCellValue(88, 7, "INPUTS\n(Hours* # of times * # of persons)");

            fmt = xls.GetCellVisibleFormatDef(89, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.Borders.Top.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Top.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(89, 3, xls.AddFormat(fmt));
            xls.SetCellValue(89, 3, "Árbol maduro");

            fmt = xls.GetCellVisibleFormatDef(89, 4);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.Borders.Top.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Top.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(89, 4, xls.AddFormat(fmt));
            xls.SetCellValue(89, 4, "Años 4, 5 y 6");

            fmt = xls.GetCellVisibleFormatDef(89, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Top.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Top.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.WrapText = true;
            xls.SetCellFormat(89, 5, xls.AddFormat(fmt));
            xls.SetCellValue(89, 5, "Mano de obra para mantenimiento");

            fmt = xls.GetCellVisibleFormatDef(89, 6);
            fmt.Borders.Top.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Top.Color = TExcelColor.Automatic;
            xls.SetCellFormat(89, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(89, 7);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.Borders.Top.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Top.Color = TExcelColor.Automatic;
            xls.SetCellFormat(89, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(90, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(90, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(90, 4);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(90, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(90, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(90, 5, xls.AddFormat(fmt));
            xls.SetCellValue(90, 5, "Desyerbe para mantenimiento");

            fmt = xls.GetCellVisibleFormatDef(90, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "0.00";
            xls.SetCellFormat(90, 6, xls.AddFormat(fmt));
            xls.SetCellValue(90, 6, 31);

            fmt = xls.GetCellVisibleFormatDef(90, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(90, 7, xls.AddFormat(fmt));
            xls.SetCellValue(90, 7, "INPUTS\n(Days * # of times * # of persons)");

            fmt = xls.GetCellVisibleFormatDef(91, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(91, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(91, 4);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(91, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(91, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(91, 5, xls.AddFormat(fmt));
            xls.SetCellValue(91, 5, "Desyerbe químico para mantenimiento");

            fmt = xls.GetCellVisibleFormatDef(91, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "0.00";
            xls.SetCellFormat(91, 6, xls.AddFormat(fmt));
            xls.SetCellValue(91, 6, 0.04);

            fmt = xls.GetCellVisibleFormatDef(91, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(91, 7, xls.AddFormat(fmt));
            xls.SetCellValue(91, 7, "INPUTS\n(Days * # of times * # of persons)");

            fmt = xls.GetCellVisibleFormatDef(92, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(92, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(92, 4);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(92, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(92, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(92, 5, xls.AddFormat(fmt));
            xls.SetCellValue(92, 5, "Aplicación de abonos orgánicos para mantenimiento");

            fmt = xls.GetCellVisibleFormatDef(92, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "0.00";
            xls.SetCellFormat(92, 6, xls.AddFormat(fmt));
            xls.SetCellValue(92, 6, 5.5);

            fmt = xls.GetCellVisibleFormatDef(92, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(92, 7, xls.AddFormat(fmt));
            xls.SetCellValue(92, 7, "INPUTS\n(Days * # of times * # of persons)");

            fmt = xls.GetCellVisibleFormatDef(93, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(93, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(93, 4);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(93, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(93, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(93, 5, xls.AddFormat(fmt));
            xls.SetCellValue(93, 5, "Aplicación de abonos químicos para mantenimiento");

            fmt = xls.GetCellVisibleFormatDef(93, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "0.00";
            xls.SetCellFormat(93, 6, xls.AddFormat(fmt));
            xls.SetCellValue(93, 6, 0);

            fmt = xls.GetCellVisibleFormatDef(93, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(93, 7, xls.AddFormat(fmt));
            xls.SetCellValue(93, 7, "INPUTS\n(Days * # of times * # of persons)");

            fmt = xls.GetCellVisibleFormatDef(94, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(94, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(94, 4);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(94, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(94, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(94, 5, xls.AddFormat(fmt));
            xls.SetCellValue(94, 5, "Aplicación de foliares para fertilización y control roya");

            fmt = xls.GetCellVisibleFormatDef(94, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "0.00";
            xls.SetCellFormat(94, 6, xls.AddFormat(fmt));
            xls.SetCellValue(94, 6, 3.4);

            fmt = xls.GetCellVisibleFormatDef(94, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(94, 7, xls.AddFormat(fmt));
            xls.SetCellValue(94, 7, "INPUTS\n(Days * # of times * # of persons)");

            fmt = xls.GetCellVisibleFormatDef(95, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(95, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(95, 4);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(95, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(95, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(95, 5, xls.AddFormat(fmt));
            xls.SetCellValue(95, 5, "Construcción de barreras vivas (rompe-vientos)");

            fmt = xls.GetCellVisibleFormatDef(95, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "0.00";
            xls.SetCellFormat(95, 6, xls.AddFormat(fmt));
            xls.SetCellValue(95, 6, 2.5);

            fmt = xls.GetCellVisibleFormatDef(95, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(95, 7, xls.AddFormat(fmt));
            xls.SetCellValue(95, 7, "INPUTS\n(Days * # of times * # of persons)");

            fmt = xls.GetCellVisibleFormatDef(96, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(96, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(96, 4);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(96, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(96, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(96, 5, xls.AddFormat(fmt));
            xls.SetCellValue(96, 5, "Podas de árboles de sombra (sostenimiento)");

            fmt = xls.GetCellVisibleFormatDef(96, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "0.00";
            xls.SetCellFormat(96, 6, xls.AddFormat(fmt));
            xls.SetCellValue(96, 6, 11.7);

            fmt = xls.GetCellVisibleFormatDef(96, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(96, 7, xls.AddFormat(fmt));
            xls.SetCellValue(96, 7, "INPUTS\n(Days * # of times * # of persons)");

            fmt = xls.GetCellVisibleFormatDef(97, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(97, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(97, 4);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(97, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(97, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.left;
            xls.SetCellFormat(97, 5, xls.AddFormat(fmt));
            xls.SetCellValue(97, 5, "Control de Broca (re-re, repela, fumigaciones)");

            fmt = xls.GetCellVisibleFormatDef(97, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "0.00";
            xls.SetCellFormat(97, 6, xls.AddFormat(fmt));
            xls.SetCellValue(97, 6, 0.36);

            fmt = xls.GetCellVisibleFormatDef(97, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(97, 7, xls.AddFormat(fmt));
            xls.SetCellValue(97, 7, "INPUTS\n(Days * # of times * # of persons)");

            fmt = xls.GetCellVisibleFormatDef(98, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(98, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(98, 4);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(98, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(98, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(98, 5, xls.AddFormat(fmt));
            xls.SetCellValue(98, 5, "Manejo de tejido (desrrame o podas del café)");

            fmt = xls.GetCellVisibleFormatDef(98, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "0.00";
            xls.SetCellFormat(98, 6, xls.AddFormat(fmt));
            xls.SetCellValue(98, 6, 3.91);

            fmt = xls.GetCellVisibleFormatDef(98, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(98, 7, xls.AddFormat(fmt));
            xls.SetCellValue(98, 7, "INPUTS\n(Days * # of times * # of persons)");

            fmt = xls.GetCellVisibleFormatDef(99, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(99, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(99, 4);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(99, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(99, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.left;
            xls.SetCellFormat(99, 5, xls.AddFormat(fmt));
            xls.SetCellValue(99, 5, "Otros:");

            fmt = xls.GetCellVisibleFormatDef(99, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "0.00";
            xls.SetCellFormat(99, 6, xls.AddFormat(fmt));
            xls.SetCellValue(99, 6, 7.36);

            fmt = xls.GetCellVisibleFormatDef(99, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(99, 7, xls.AddFormat(fmt));
            xls.SetCellValue(99, 7, "INPUTS\n(Days * # of times * # of persons)");

            fmt = xls.GetCellVisibleFormatDef(100, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(100, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(100, 4);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(100, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(100, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.WrapText = true;
            xls.SetCellFormat(100, 5, xls.AddFormat(fmt));
            xls.SetCellValue(100, 5, "Mano de obra para cosecha");

            fmt = xls.GetCellVisibleFormatDef(100, 7);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            xls.SetCellFormat(100, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(101, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(101, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(101, 4);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(101, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(101, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Locked = false;
            fmt.WrapText = true;
            xls.SetCellFormat(101, 5, xls.AddFormat(fmt));
            xls.SetCellValue(101, 5, "Cuantos Quintales cosecha por hectarea");
            xls.SetCellValue(101, 6, 10);

            fmt = xls.GetCellVisibleFormatDef(101, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(101, 7, xls.AddFormat(fmt));
            xls.SetCellValue(101, 7, "INPUTS\n#");

            fmt = xls.GetCellVisibleFormatDef(102, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(102, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(102, 4);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(102, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(102, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(102, 5, xls.AddFormat(fmt));
            xls.SetCellValue(102, 5, "TOTAL de tiempo recogiendo café");
            xls.SetCellValue(102, 6, 65);

            fmt = xls.GetCellVisibleFormatDef(102, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(102, 7, xls.AddFormat(fmt));
            xls.SetCellValue(102, 7, "INPUTS\n(Days * # of times * # of persons)");

            fmt = xls.GetCellVisibleFormatDef(103, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(103, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(103, 4);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(103, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(103, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(103, 5, xls.AddFormat(fmt));
            xls.SetCellValue(103, 5, "Otras actividades relacionadas con la cosecha");
            xls.SetCellValue(103, 6, 0);

            fmt = xls.GetCellVisibleFormatDef(103, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(103, 7, xls.AddFormat(fmt));
            xls.SetCellValue(103, 7, "INPUTS\n#");

            fmt = xls.GetCellVisibleFormatDef(104, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(104, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(104, 4);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(104, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(104, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(104, 5, xls.AddFormat(fmt));
            xls.SetCellValue(104, 5, "¿Cuántos KILOS de CEREZA recoge POR HECTÁREA?");

            fmt = xls.GetCellVisibleFormatDef(104, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "0.00";
            xls.SetCellFormat(104, 6, xls.AddFormat(fmt));
            xls.SetCellValue(104, 6, 10);

            fmt = xls.GetCellVisibleFormatDef(104, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(104, 7, xls.AddFormat(fmt));
            xls.SetCellValue(104, 7, "INPUTS\n#");

            fmt = xls.GetCellVisibleFormatDef(105, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(105, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(105, 4);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(105, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(105, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.left;
            xls.SetCellFormat(105, 5, xls.AddFormat(fmt));

            Runs = new TRTFRun[2];
            Runs[0].FirstChar = 9;
            fnt = xls.GetDefaultFont;
            fnt.Name = "Arial";
            fnt.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fnt.Style = TFlxFontStyles.Bold;
            fnt.Scheme = TFontScheme.None;
            Runs[0].FontIndex = xls.AddFont(fnt);
            Runs[1].FirstChar = 18;
            fnt = xls.GetDefaultFont;
            fnt.Name = "Arial";
            fnt.Color = TExcelColor.Automatic;
            fnt.Scheme = TFontScheme.None;
            Runs[1].FontIndex = xls.AddFont(fnt);
            xls.SetCellValue(105, 5, new TRichString("¿Cuántas QUINTALES de PERGAMINO SECO recoge POR HECTÁREA?", Runs, xls));
            //We could also have used: xls.SetCellFromHtml(105, 5, "&iquest;Cu&aacute;ntas&nbsp;<font color = 'red'><b>QUINTALES</b></font>&nbsp;de PERGAMINO"
            //+" SECO recoge POR HECT&Aacute;REA?")


    fmt = xls.GetCellVisibleFormatDef(105, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "0.00";
            xls.SetCellFormat(105, 6, xls.AddFormat(fmt));
            xls.SetCellValue(105, 6, new TFormula("=Proportions!J6"));

            fmt = xls.GetCellVisibleFormatDef(105, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(105, 7, xls.AddFormat(fmt));
            xls.SetCellValue(105, 7, "INPUTS\n#");

            fmt = xls.GetCellVisibleFormatDef(106, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(106, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(106, 4);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(106, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(106, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.WrapText = true;
            xls.SetCellFormat(106, 5, xls.AddFormat(fmt));
            xls.SetCellValue(106, 5, "Mano de obra para el beneficio (en Horas)");

            fmt = xls.GetCellVisibleFormatDef(106, 7);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            xls.SetCellFormat(106, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(107, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(107, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(107, 4);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(107, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(107, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(107, 5, xls.AddFormat(fmt));
            xls.SetCellValue(107, 5, "Despulpado y Fermentado");

            fmt = xls.GetCellVisibleFormatDef(107, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "0.00";
            xls.SetCellFormat(107, 6, xls.AddFormat(fmt));
            xls.SetCellValue(107, 6, 6.5);

            fmt = xls.GetCellVisibleFormatDef(107, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(107, 7, xls.AddFormat(fmt));
            xls.SetCellValue(107, 7, "INPUTS\n(Hours * # of times * # of persons)");

            fmt = xls.GetCellVisibleFormatDef(108, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(108, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(108, 4);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(108, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(108, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(108, 5, xls.AddFormat(fmt));
            xls.SetCellValue(108, 5, "Lavado (incluye rebalse)");

            fmt = xls.GetCellVisibleFormatDef(108, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "0.00";
            xls.SetCellFormat(108, 6, xls.AddFormat(fmt));
            xls.SetCellValue(108, 6, 6);

            fmt = xls.GetCellVisibleFormatDef(108, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(108, 7, xls.AddFormat(fmt));
            xls.SetCellValue(108, 7, "INPUTS\n(Hours * # of times * # of persons)");

            fmt = xls.GetCellVisibleFormatDef(109, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(109, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(109, 4);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(109, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(109, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(109, 5, xls.AddFormat(fmt));
            xls.SetCellValue(109, 5, "Secado");

            fmt = xls.GetCellVisibleFormatDef(109, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "0.00";
            xls.SetCellFormat(109, 6, xls.AddFormat(fmt));
            xls.SetCellValue(109, 6, 8.5);

            fmt = xls.GetCellVisibleFormatDef(109, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(109, 7, xls.AddFormat(fmt));
            xls.SetCellValue(109, 7, "INPUTS\n(Hours * # of times * # of persons)");

            fmt = xls.GetCellVisibleFormatDef(110, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(110, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(110, 4);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(110, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(110, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(110, 5, xls.AddFormat(fmt));
            xls.SetCellValue(110, 5, "Zarandeo");

            fmt = xls.GetCellVisibleFormatDef(110, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "0.00";
            xls.SetCellFormat(110, 6, xls.AddFormat(fmt));
            xls.SetCellValue(110, 6, 2.13);

            fmt = xls.GetCellVisibleFormatDef(110, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(110, 7, xls.AddFormat(fmt));
            xls.SetCellValue(110, 7, "INPUTS\n(Hours * # of times * # of persons)");

            fmt = xls.GetCellVisibleFormatDef(111, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(111, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(111, 4);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(111, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(111, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.left;
            xls.SetCellFormat(111, 5, xls.AddFormat(fmt));
            xls.SetCellValue(111, 5, "Escojo/selección");

            fmt = xls.GetCellVisibleFormatDef(111, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "0.00";
            xls.SetCellFormat(111, 6, xls.AddFormat(fmt));
            xls.SetCellValue(111, 6, 4.8);

            fmt = xls.GetCellVisibleFormatDef(111, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(111, 7, xls.AddFormat(fmt));
            xls.SetCellValue(111, 7, "INPUTS\n(Hours * # of times * # of persons)");

            fmt = xls.GetCellVisibleFormatDef(112, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(112, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(112, 4);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(112, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(112, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(112, 5, xls.AddFormat(fmt));
            xls.SetCellValue(112, 5, "Almacenamiento");

            fmt = xls.GetCellVisibleFormatDef(112, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "0.00";
            xls.SetCellFormat(112, 6, xls.AddFormat(fmt));
            xls.SetCellValue(112, 6, 2.3);

            fmt = xls.GetCellVisibleFormatDef(112, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(112, 7, xls.AddFormat(fmt));
            xls.SetCellValue(112, 7, "INPUTS\n(Hours * # of times * # of persons)");

            fmt = xls.GetCellVisibleFormatDef(113, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(113, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(113, 4);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(113, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(113, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(113, 5, xls.AddFormat(fmt));
            xls.SetCellValue(113, 5, "Aguas Miel");

            fmt = xls.GetCellVisibleFormatDef(113, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "0.00";
            xls.SetCellFormat(113, 6, xls.AddFormat(fmt));
            xls.SetCellValue(113, 6, 0.43);

            fmt = xls.GetCellVisibleFormatDef(113, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(113, 7, xls.AddFormat(fmt));
            xls.SetCellValue(113, 7, "INPUTS\n(Hours * # of times * # of persons)");

            fmt = xls.GetCellVisibleFormatDef(114, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(114, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(114, 4);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(114, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(114, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(114, 5, xls.AddFormat(fmt));
            xls.SetCellValue(114, 5, "Manejo de pulpa");

            fmt = xls.GetCellVisibleFormatDef(114, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "0.00";
            xls.SetCellFormat(114, 6, xls.AddFormat(fmt));
            xls.SetCellValue(114, 6, 3);

            fmt = xls.GetCellVisibleFormatDef(114, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(114, 7, xls.AddFormat(fmt));
            xls.SetCellValue(114, 7, "INPUTS\n(Hours * # of times * # of persons)");

            fmt = xls.GetCellVisibleFormatDef(115, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(115, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(115, 4);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(115, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(115, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            xls.SetCellFormat(115, 5, xls.AddFormat(fmt));
            xls.SetCellValue(115, 5, "Otros");

            fmt = xls.GetCellVisibleFormatDef(115, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.Format = "0.00";
            xls.SetCellFormat(115, 6, xls.AddFormat(fmt));
            xls.SetCellValue(115, 6, 0.1);

            fmt = xls.GetCellVisibleFormatDef(115, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(115, 7, xls.AddFormat(fmt));
            xls.SetCellValue(115, 7, "INPUTS\n(Hours * # of times * # of persons)");

            fmt = xls.GetCellVisibleFormatDef(116, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.Borders.Top.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Top.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(116, 3, xls.AddFormat(fmt));
            xls.SetCellValue(116, 3, "Árbol viejo");

            fmt = xls.GetCellVisibleFormatDef(116, 4);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.Borders.Top.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Top.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(116, 4, xls.AddFormat(fmt));
            xls.SetCellValue(116, 4, "Áños 7 y 8");

            fmt = xls.GetCellVisibleFormatDef(116, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Top.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Top.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.WrapText = true;
            xls.SetCellFormat(116, 5, xls.AddFormat(fmt));
            xls.SetCellValue(116, 5, "Mano de obra para mantenimiento");

            fmt = xls.GetCellVisibleFormatDef(116, 6);
            fmt.Borders.Top.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Top.Color = TExcelColor.Automatic;
            xls.SetCellFormat(116, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(116, 7);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.Borders.Top.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Top.Color = TExcelColor.Automatic;
            xls.SetCellFormat(116, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(117, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(117, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(117, 4);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(117, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(117, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(117, 5, xls.AddFormat(fmt));
            xls.SetCellValue(117, 5, "Desyerbe para mantenimiento");

            fmt = xls.GetCellVisibleFormatDef(117, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "0.00";
            xls.SetCellFormat(117, 6, xls.AddFormat(fmt));
            xls.SetCellValue(117, 6, 28);

            fmt = xls.GetCellVisibleFormatDef(117, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(117, 7, xls.AddFormat(fmt));
            xls.SetCellValue(117, 7, "INPUTS\n(Days * # of times * # of persons)");

            fmt = xls.GetCellVisibleFormatDef(118, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(118, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(118, 4);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(118, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(118, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(118, 5, xls.AddFormat(fmt));
            xls.SetCellValue(118, 5, "Desyerbe químico para mantenimiento");

            fmt = xls.GetCellVisibleFormatDef(118, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "0.00";
            xls.SetCellFormat(118, 6, xls.AddFormat(fmt));
            xls.SetCellValue(118, 6, 0.04);

            fmt = xls.GetCellVisibleFormatDef(118, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(118, 7, xls.AddFormat(fmt));
            xls.SetCellValue(118, 7, "INPUTS\n(Days * # of times * # of persons)");

            fmt = xls.GetCellVisibleFormatDef(119, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(119, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(119, 4);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(119, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(119, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(119, 5, xls.AddFormat(fmt));
            xls.SetCellValue(119, 5, "Aplicación de abonos orgánicos para mantenimiento");

            fmt = xls.GetCellVisibleFormatDef(119, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "0.00";
            xls.SetCellFormat(119, 6, xls.AddFormat(fmt));
            xls.SetCellValue(119, 6, 5.78);

            fmt = xls.GetCellVisibleFormatDef(119, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(119, 7, xls.AddFormat(fmt));
            xls.SetCellValue(119, 7, "INPUTS\n(Days * # of times * # of persons)");

            fmt = xls.GetCellVisibleFormatDef(120, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(120, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(120, 4);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(120, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(120, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(120, 5, xls.AddFormat(fmt));
            xls.SetCellValue(120, 5, "Aplicación de abonos químicos para mantenimiento");

            fmt = xls.GetCellVisibleFormatDef(120, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "0.00";
            xls.SetCellFormat(120, 6, xls.AddFormat(fmt));
            xls.SetCellValue(120, 6, 0);

            fmt = xls.GetCellVisibleFormatDef(120, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(120, 7, xls.AddFormat(fmt));
            xls.SetCellValue(120, 7, "INPUTS\n(Days * # of times * # of persons)");

            fmt = xls.GetCellVisibleFormatDef(121, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(121, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(121, 4);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(121, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(121, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(121, 5, xls.AddFormat(fmt));
            xls.SetCellValue(121, 5, "Aplicación de foliares para fertilización y control roya");

            fmt = xls.GetCellVisibleFormatDef(121, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "0.00";
            xls.SetCellFormat(121, 6, xls.AddFormat(fmt));
            xls.SetCellValue(121, 6, 3.71);

            fmt = xls.GetCellVisibleFormatDef(121, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(121, 7, xls.AddFormat(fmt));
            xls.SetCellValue(121, 7, "INPUTS\n(Days * # of times * # of persons)");

            fmt = xls.GetCellVisibleFormatDef(122, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(122, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(122, 4);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(122, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(122, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(122, 5, xls.AddFormat(fmt));
            xls.SetCellValue(122, 5, "Construcción de barreras vivas (rompe-vientos)");

            fmt = xls.GetCellVisibleFormatDef(122, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "0.00";
            xls.SetCellFormat(122, 6, xls.AddFormat(fmt));
            xls.SetCellValue(122, 6, 2.2);

            fmt = xls.GetCellVisibleFormatDef(122, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(122, 7, xls.AddFormat(fmt));
            xls.SetCellValue(122, 7, "INPUTS\n(Days * # of times * # of persons)");

            fmt = xls.GetCellVisibleFormatDef(123, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(123, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(123, 4);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(123, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(123, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(123, 5, xls.AddFormat(fmt));
            xls.SetCellValue(123, 5, "Podas de árboles de sombra (sostenimiento)");

            fmt = xls.GetCellVisibleFormatDef(123, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "0.00";
            xls.SetCellFormat(123, 6, xls.AddFormat(fmt));
            xls.SetCellValue(123, 6, 12.2);

            fmt = xls.GetCellVisibleFormatDef(123, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(123, 7, xls.AddFormat(fmt));
            xls.SetCellValue(123, 7, "INPUTS\n(Days * # of times * # of persons)");

            fmt = xls.GetCellVisibleFormatDef(124, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(124, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(124, 4);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(124, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(124, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.left;
            xls.SetCellFormat(124, 5, xls.AddFormat(fmt));
            xls.SetCellValue(124, 5, "Control de Broca (re-re, repela, fumigaciones)");

            fmt = xls.GetCellVisibleFormatDef(124, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "0.00";
            xls.SetCellFormat(124, 6, xls.AddFormat(fmt));
            xls.SetCellValue(124, 6, 0.36);

            fmt = xls.GetCellVisibleFormatDef(124, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(124, 7, xls.AddFormat(fmt));
            xls.SetCellValue(124, 7, "INPUTS\n(Days * # of times * # of persons)");

            fmt = xls.GetCellVisibleFormatDef(125, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(125, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(125, 4);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(125, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(125, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(125, 5, xls.AddFormat(fmt));
            xls.SetCellValue(125, 5, "Manejo de tejido (desrrame o podas del café)");

            fmt = xls.GetCellVisibleFormatDef(125, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "0.00";
            xls.SetCellFormat(125, 6, xls.AddFormat(fmt));
            xls.SetCellValue(125, 6, 4.54);

            fmt = xls.GetCellVisibleFormatDef(125, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(125, 7, xls.AddFormat(fmt));
            xls.SetCellValue(125, 7, "INPUTS\n(Days * # of times * # of persons)");

            fmt = xls.GetCellVisibleFormatDef(126, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(126, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(126, 4);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(126, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(126, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.left;
            xls.SetCellFormat(126, 5, xls.AddFormat(fmt));
            xls.SetCellValue(126, 5, "Otros:");

            fmt = xls.GetCellVisibleFormatDef(126, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "0.00";
            xls.SetCellFormat(126, 6, xls.AddFormat(fmt));
            xls.SetCellValue(126, 6, 7.91);

            fmt = xls.GetCellVisibleFormatDef(126, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(126, 7, xls.AddFormat(fmt));
            xls.SetCellValue(126, 7, "INPUTS\n(Days * # of times * # of persons)");

            fmt = xls.GetCellVisibleFormatDef(127, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(127, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(127, 4);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(127, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(127, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.WrapText = true;
            xls.SetCellFormat(127, 5, xls.AddFormat(fmt));
            xls.SetCellValue(127, 5, "Mano de obra para cosecha");

            fmt = xls.GetCellVisibleFormatDef(127, 7);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            xls.SetCellFormat(127, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(128, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(128, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(128, 4);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(128, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(128, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Locked = false;
            fmt.WrapText = true;
            xls.SetCellFormat(128, 5, xls.AddFormat(fmt));
            xls.SetCellValue(128, 5, "Cuantos Quintales cosecha por hectarea");

            fmt = xls.GetCellVisibleFormatDef(128, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "0.00";
            xls.SetCellFormat(128, 6, xls.AddFormat(fmt));
            xls.SetCellValue(128, 6, 16);

            fmt = xls.GetCellVisibleFormatDef(128, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(128, 7, xls.AddFormat(fmt));
            xls.SetCellValue(128, 7, "INPUTS\n#");

            fmt = xls.GetCellVisibleFormatDef(129, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(129, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(129, 4);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(129, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(129, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(129, 5, xls.AddFormat(fmt));
            xls.SetCellValue(129, 5, "TOTAL de tiempo recogiendo café");
            xls.SetCellValue(129, 6, 53);

            fmt = xls.GetCellVisibleFormatDef(129, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(129, 7, xls.AddFormat(fmt));
            xls.SetCellValue(129, 7, "INPUTS\n(Days * # of times * # of persons)");

            fmt = xls.GetCellVisibleFormatDef(130, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(130, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(130, 4);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(130, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(130, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(130, 5, xls.AddFormat(fmt));
            xls.SetCellValue(130, 5, "Otras actividades relacionadas con la cosecha");

            fmt = xls.GetCellVisibleFormatDef(130, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "0.00";
            xls.SetCellFormat(130, 6, xls.AddFormat(fmt));
            xls.SetCellValue(130, 6, 0);

            fmt = xls.GetCellVisibleFormatDef(130, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(130, 7, xls.AddFormat(fmt));
            xls.SetCellValue(130, 7, "INPUTS\n#");

            fmt = xls.GetCellVisibleFormatDef(131, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(131, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(131, 4);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(131, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(131, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(131, 5, xls.AddFormat(fmt));
            xls.SetCellValue(131, 5, "¿Cuántos KILOS de CEREZA recoge POR HECTÁREA?");

            fmt = xls.GetCellVisibleFormatDef(131, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "0.00";
            xls.SetCellFormat(131, 6, xls.AddFormat(fmt));
            xls.SetCellValue(131, 6, 16);

            fmt = xls.GetCellVisibleFormatDef(131, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(131, 7, xls.AddFormat(fmt));
            xls.SetCellValue(131, 7, "INPUTS\n#");

            fmt = xls.GetCellVisibleFormatDef(132, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(132, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(132, 4);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(132, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(132, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.left;
            xls.SetCellFormat(132, 5, xls.AddFormat(fmt));

            Runs = new TRTFRun[2];
            Runs[0].FirstChar = 9;
            fnt = xls.GetDefaultFont;
            fnt.Name = "Arial";
            fnt.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fnt.Style = TFlxFontStyles.Bold;
            fnt.Scheme = TFontScheme.None;
            Runs[0].FontIndex = xls.AddFont(fnt);
            Runs[1].FirstChar = 18;
            fnt = xls.GetDefaultFont;
            fnt.Name = "Arial";
            fnt.Color = TExcelColor.Automatic;
            fnt.Scheme = TFontScheme.None;
            Runs[1].FontIndex = xls.AddFont(fnt);
            xls.SetCellValue(132, 5, new TRichString("¿Cuántas QUINTALES de PERGAMINO SECO recoge POR HECTÁREA?", Runs, xls));
            //We could also have used: xls.SetCellFromHtml(132, 5, "&iquest;Cu&aacute;ntas&nbsp;<font color = 'red'><b>QUINTALES</b></font>&nbsp;de PERGAMINO"
            //+" SECO recoge POR HECT&Aacute;REA?")


    fmt = xls.GetCellVisibleFormatDef(132, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "0.00";
            xls.SetCellFormat(132, 6, xls.AddFormat(fmt));
            xls.SetCellValue(132, 6, new TFormula("=Proportions!J7"));

            fmt = xls.GetCellVisibleFormatDef(132, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(132, 7, xls.AddFormat(fmt));
            xls.SetCellValue(132, 7, "INPUTS\n#");

            fmt = xls.GetCellVisibleFormatDef(133, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(133, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(133, 4);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(133, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(133, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.WrapText = true;
            xls.SetCellFormat(133, 5, xls.AddFormat(fmt));
            xls.SetCellValue(133, 5, "Mano de obra para el beneficio (en Horas)");

            fmt = xls.GetCellVisibleFormatDef(133, 7);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            xls.SetCellFormat(133, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(134, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(134, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(134, 4);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(134, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(134, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(134, 5, xls.AddFormat(fmt));
            xls.SetCellValue(134, 5, "Despulpado y Fermentado");

            fmt = xls.GetCellVisibleFormatDef(134, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "0.00";
            xls.SetCellFormat(134, 6, xls.AddFormat(fmt));
            xls.SetCellValue(134, 6, 4.6);

            fmt = xls.GetCellVisibleFormatDef(134, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(134, 7, xls.AddFormat(fmt));
            xls.SetCellValue(134, 7, "INPUTS\n(Hours * # of times * # of persons)");

            fmt = xls.GetCellVisibleFormatDef(135, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(135, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(135, 4);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(135, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(135, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(135, 5, xls.AddFormat(fmt));
            xls.SetCellValue(135, 5, "Lavado (incluye rebalse)");

            fmt = xls.GetCellVisibleFormatDef(135, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "0.00";
            xls.SetCellFormat(135, 6, xls.AddFormat(fmt));
            xls.SetCellValue(135, 6, 2.3);

            fmt = xls.GetCellVisibleFormatDef(135, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(135, 7, xls.AddFormat(fmt));
            xls.SetCellValue(135, 7, "INPUTS\n(Hours * # of times * # of persons)");

            fmt = xls.GetCellVisibleFormatDef(136, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(136, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(136, 4);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(136, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(136, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(136, 5, xls.AddFormat(fmt));
            xls.SetCellValue(136, 5, "Secado");

            fmt = xls.GetCellVisibleFormatDef(136, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "0.00";
            xls.SetCellFormat(136, 6, xls.AddFormat(fmt));
            xls.SetCellValue(136, 6, 1.2);

            fmt = xls.GetCellVisibleFormatDef(136, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(136, 7, xls.AddFormat(fmt));
            xls.SetCellValue(136, 7, "INPUTS\n(Hours * # of times * # of persons)");

            fmt = xls.GetCellVisibleFormatDef(137, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(137, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(137, 4);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(137, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(137, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(137, 5, xls.AddFormat(fmt));
            xls.SetCellValue(137, 5, "Zarandeo");

            fmt = xls.GetCellVisibleFormatDef(137, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "0.00";
            xls.SetCellFormat(137, 6, xls.AddFormat(fmt));
            xls.SetCellValue(137, 6, 0.83);

            fmt = xls.GetCellVisibleFormatDef(137, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(137, 7, xls.AddFormat(fmt));
            xls.SetCellValue(137, 7, "INPUTS\n(Hours * # of times * # of persons)");

            fmt = xls.GetCellVisibleFormatDef(138, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(138, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(138, 4);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(138, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(138, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.left;
            xls.SetCellFormat(138, 5, xls.AddFormat(fmt));
            xls.SetCellValue(138, 5, "Escojo/selección");

            fmt = xls.GetCellVisibleFormatDef(138, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "0.00";
            xls.SetCellFormat(138, 6, xls.AddFormat(fmt));
            xls.SetCellValue(138, 6, 0);

            fmt = xls.GetCellVisibleFormatDef(138, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(138, 7, xls.AddFormat(fmt));
            xls.SetCellValue(138, 7, "INPUTS\n(Hours * # of times * # of persons)");

            fmt = xls.GetCellVisibleFormatDef(139, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(139, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(139, 4);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(139, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(139, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(139, 5, xls.AddFormat(fmt));
            xls.SetCellValue(139, 5, "Almacenamiento");

            fmt = xls.GetCellVisibleFormatDef(139, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "0.00";
            xls.SetCellFormat(139, 6, xls.AddFormat(fmt));
            xls.SetCellValue(139, 6, 0.21);

            fmt = xls.GetCellVisibleFormatDef(139, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(139, 7, xls.AddFormat(fmt));
            xls.SetCellValue(139, 7, "INPUTS\n(Hours * # of times * # of persons)");

            fmt = xls.GetCellVisibleFormatDef(140, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(140, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(140, 4);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(140, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(140, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(140, 5, xls.AddFormat(fmt));
            xls.SetCellValue(140, 5, "Aguas Miel");

            fmt = xls.GetCellVisibleFormatDef(140, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "0.00";
            xls.SetCellFormat(140, 6, xls.AddFormat(fmt));
            xls.SetCellValue(140, 6, 0);

            fmt = xls.GetCellVisibleFormatDef(140, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(140, 7, xls.AddFormat(fmt));
            xls.SetCellValue(140, 7, "INPUTS\n(Hours * # of times * # of persons)");

            fmt = xls.GetCellVisibleFormatDef(141, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(141, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(141, 4);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(141, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(141, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(141, 5, xls.AddFormat(fmt));
            xls.SetCellValue(141, 5, "Manejo de pulpa");

            fmt = xls.GetCellVisibleFormatDef(141, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "0.00";
            xls.SetCellFormat(141, 6, xls.AddFormat(fmt));
            xls.SetCellValue(141, 6, 0.7);

            fmt = xls.GetCellVisibleFormatDef(141, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(141, 7, xls.AddFormat(fmt));
            xls.SetCellValue(141, 7, "INPUTS\n(Hours * # of times * # of persons)");

            fmt = xls.GetCellVisibleFormatDef(142, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(142, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(142, 4);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(142, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(142, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            xls.SetCellFormat(142, 5, xls.AddFormat(fmt));
            xls.SetCellValue(142, 5, "Otros");

            fmt = xls.GetCellVisibleFormatDef(142, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.Format = "0.00";
            xls.SetCellFormat(142, 6, xls.AddFormat(fmt));
            xls.SetCellValue(142, 6, 0);

            fmt = xls.GetCellVisibleFormatDef(142, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(142, 7, xls.AddFormat(fmt));
            xls.SetCellValue(142, 7, "INPUTS\n(Hours * # of times * # of persons)");

            fmt = xls.GetCellVisibleFormatDef(143, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.Borders.Top.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Top.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(143, 3, xls.AddFormat(fmt));
            xls.SetCellValue(143, 3, "Información general");

            fmt = xls.GetCellVisibleFormatDef(143, 4);
            fmt.Borders.Top.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Top.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(143, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(143, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Top.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Top.Color = TExcelColor.Automatic;
            xls.SetCellFormat(143, 5, xls.AddFormat(fmt));
            xls.SetCellValue(143, 5, "¿Cuántos KILOS de cereza recoge en promedio una persona en UN DÍA?");

            fmt = xls.GetCellVisibleFormatDef(143, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Top.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Top.Color = TExcelColor.Automatic;
            fmt.Format = "0.00";
            xls.SetCellFormat(143, 6, xls.AddFormat(fmt));
            xls.SetCellValue(143, 6, 1.5);

            fmt = xls.GetCellVisibleFormatDef(143, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.Borders.Top.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Top.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(143, 7, xls.AddFormat(fmt));
            xls.SetCellValue(143, 7, "INPUTS\n#");

            fmt = xls.GetCellVisibleFormatDef(144, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(144, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(144, 4);
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(144, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(144, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Locked = false;
            fmt.WrapText = true;
            xls.SetCellFormat(144, 5, xls.AddFormat(fmt));
            xls.SetCellValue(144, 5, "Cuanto paga por Caja Recolectada");

            fmt = xls.GetCellVisibleFormatDef(144, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "0.00";
            xls.SetCellFormat(144, 6, xls.AddFormat(fmt));
            xls.SetCellValue(144, 6, 80);

            fmt = xls.GetCellVisibleFormatDef(144, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(144, 7, xls.AddFormat(fmt));
            xls.SetCellValue(144, 7, "INPUTS\n#");

            fmt = xls.GetCellVisibleFormatDef(145, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(145, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(145, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 320;
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(145, 5, xls.AddFormat(fmt));
            xls.SetCellValue(145, 5, "Informacion General");

            fmt = xls.GetCellVisibleFormatDef(145, 7);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            xls.SetCellFormat(145, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(146, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(146, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(146, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.WrapText = true;
            xls.SetCellFormat(146, 5, xls.AddFormat(fmt));

            Runs = new TRTFRun[1];
            Runs[0].FirstChar = 84;
            fnt = xls.GetDefaultFont;
            fnt.Name = "Arial";
            fnt.Size20 = 280;
            fnt.Style = TFlxFontStyles.Italic;
            fnt.Scheme = TFontScheme.None;
            Runs[0].FontIndex = xls.AddFont(fnt);
            xls.SetCellValue(146, 5, new TRichString("¿Cuánto tiempo lleva usted en la actividad cafetera?                             "
            + "   *Nota: Si toda la vida, poner un estimado de los años.", Runs, xls));
            //We could also have used: xls.SetCellFromHtml(146, 5, "&iquest;Cu&aacute;nto tiempo lleva usted en la actividad cafetera? &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<i>*Nota:"
            //+" Si toda la vida, poner un estimado de los a&ntilde;os.</i>")


    fmt = xls.GetCellVisibleFormatDef(146, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(146, 6, xls.AddFormat(fmt));
            xls.SetCellValue(146, 6, 25.8);

            fmt = xls.GetCellVisibleFormatDef(146, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(146, 7, xls.AddFormat(fmt));
            xls.SetCellValue(146, 7, "INPUTS\n#");

            fmt = xls.GetCellVisibleFormatDef(147, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(147, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(147, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(147, 5, xls.AddFormat(fmt));
            xls.SetCellValue(147, 5, "¿En que año realizó  la última renovación?");

            fmt = xls.GetCellVisibleFormatDef(147, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(147, 6, xls.AddFormat(fmt));
            xls.SetCellValue(147, 6, 2015);

            fmt = xls.GetCellVisibleFormatDef(147, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(147, 7, xls.AddFormat(fmt));
            xls.SetCellValue(147, 7, "INPUTS\n#");

            fmt = xls.GetCellVisibleFormatDef(148, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(148, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(148, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 320;
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(148, 5, xls.AddFormat(fmt));
            xls.SetCellValue(148, 5, "Finca Cafetera");

            fmt = xls.GetCellVisibleFormatDef(148, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(148, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(148, 7);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            xls.SetCellFormat(148, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(149, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(149, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(149, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(149, 5, xls.AddFormat(fmt));
            xls.SetCellValue(149, 5, "Cuántas plantas de café estima en toda su finca");

            fmt = xls.GetCellVisibleFormatDef(149, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(149, 6, xls.AddFormat(fmt));
            xls.SetCellValue(149, 6, 16750);

            fmt = xls.GetCellVisibleFormatDef(149, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(149, 7, xls.AddFormat(fmt));
            xls.SetCellValue(149, 7, "INPUTS\n#");

            fmt = xls.GetCellVisibleFormatDef(150, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(150, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(150, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(150, 5, xls.AddFormat(fmt));
            xls.SetCellValue(150, 5, "Cuántas plantas de café puede sembrar en una hectárea");

            fmt = xls.GetCellVisibleFormatDef(150, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(150, 6, xls.AddFormat(fmt));
            xls.SetCellValue(150, 6, 3365);

            fmt = xls.GetCellVisibleFormatDef(150, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(150, 7, xls.AddFormat(fmt));
            xls.SetCellValue(150, 7, "INPUTS\n#");

            fmt = xls.GetCellVisibleFormatDef(151, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(151, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(151, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(151, 5, xls.AddFormat(fmt));
            xls.SetCellValue(151, 5, "Área total de la finca (hectáreas)");

            fmt = xls.GetCellVisibleFormatDef(151, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(151, 6, xls.AddFormat(fmt));
            xls.SetCellValue(151, 6, 7.87);

            fmt = xls.GetCellVisibleFormatDef(151, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(151, 7, xls.AddFormat(fmt));
            xls.SetCellValue(151, 7, "INPUTS\n#");

            fmt = xls.GetCellVisibleFormatDef(152, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(152, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(152, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(152, 5, xls.AddFormat(fmt));
            xls.SetCellValue(152, 5, "Área total en café (hectáreas)");

            fmt = xls.GetCellVisibleFormatDef(152, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent4, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(152, 6, xls.AddFormat(fmt));
            xls.SetCellValue(152, 6, new TFormula("=F153+F154+F155"));

            fmt = xls.GetCellVisibleFormatDef(152, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(152, 7, xls.AddFormat(fmt));
            xls.SetCellValue(152, 7, "INPUTS\n(hectares young trees + hectares grown up trees + hectares old trees)");

            fmt = xls.GetCellVisibleFormatDef(153, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(153, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(153, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(153, 5, xls.AddFormat(fmt));
            xls.SetCellValue(153, 5, "Indique cuántas hectáreas de café tienen árboles jóvenes");

            fmt = xls.GetCellVisibleFormatDef(153, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent4, 0.399975585192419);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(153, 6, xls.AddFormat(fmt));
            xls.SetCellValue(153, 6, new TFormula("='Inputs 1.0_metric_currency'!D6"));

            fmt = xls.GetCellVisibleFormatDef(153, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(153, 7, xls.AddFormat(fmt));
            xls.SetCellValue(153, 7, "INPUTS\n#");

            fmt = xls.GetCellVisibleFormatDef(153, 18);
            fmt.Format = "0.00";
            xls.SetCellFormat(153, 18, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(154, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(154, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(154, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(154, 5, xls.AddFormat(fmt));
            xls.SetCellValue(154, 5, "Indique cuántas hectáreas de café tienen árboles maduros");

            fmt = xls.GetCellVisibleFormatDef(154, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent4, 0.399975585192419);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(154, 6, xls.AddFormat(fmt));
            xls.SetCellValue(154, 6, new TFormula("='Inputs 1.0_metric_currency'!D7"));

            fmt = xls.GetCellVisibleFormatDef(154, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(154, 7, xls.AddFormat(fmt));
            xls.SetCellValue(154, 7, "INPUTS\n#");

            fmt = xls.GetCellVisibleFormatDef(154, 18);
            fmt.Format = "0.00";
            xls.SetCellFormat(154, 18, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(155, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(155, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(155, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(155, 5, xls.AddFormat(fmt));
            xls.SetCellValue(155, 5, "Indique cuántas hectáreas de café tienen árboles viejos");

            fmt = xls.GetCellVisibleFormatDef(155, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent4, 0.399975585192419);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(155, 6, xls.AddFormat(fmt));
            xls.SetCellValue(155, 6, new TFormula("='Inputs 1.0_metric_currency'!D8"));

            fmt = xls.GetCellVisibleFormatDef(155, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(155, 7, xls.AddFormat(fmt));
            xls.SetCellValue(155, 7, "INPUTS\n#");

            fmt = xls.GetCellVisibleFormatDef(155, 13);
            fmt.Format = "0.00";
            xls.SetCellFormat(155, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(155, 18);
            fmt.Format = "0.00";
            xls.SetCellFormat(155, 18, xls.AddFormat(fmt));
            xls.SetCellValue(156, 1, 899);

            fmt = xls.GetCellVisibleFormatDef(156, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(156, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(156, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(156, 5, xls.AddFormat(fmt));
            xls.SetCellValue(156, 5, "¿Qué variedades de café tiene en su parcela?");

            fmt = xls.GetCellVisibleFormatDef(156, 7);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            xls.SetCellFormat(156, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(156, 13);
            fmt.Format = "0.00";
            xls.SetCellFormat(156, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(156, 15);
            fmt.Format = "0.00";
            xls.SetCellFormat(156, 15, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(157, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(157, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(157, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.HAlignment = THFlxAlignment.left;
            xls.SetCellFormat(157, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(157, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.Indent = 3;
            xls.SetCellFormat(157, 5, xls.AddFormat(fmt));
            xls.SetCellValue(157, 5, "Árabe");

            fmt = xls.GetCellVisibleFormatDef(157, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(157, 6, xls.AddFormat(fmt));
            xls.SetCellValue(157, 6, 0.45);

            fmt = xls.GetCellVisibleFormatDef(157, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(157, 7, xls.AddFormat(fmt));
            xls.SetCellValue(157, 7, "INPUTS\n%");

            fmt = xls.GetCellVisibleFormatDef(157, 8);
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(157, 8, xls.AddFormat(fmt));
            xls.SetCellValue(157, 8, "Sum = 100");

            fmt = xls.GetCellVisibleFormatDef(157, 13);
            fmt.Format = "0.00";
            xls.SetCellFormat(157, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(158, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(158, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(158, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.HAlignment = THFlxAlignment.left;
            xls.SetCellFormat(158, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(158, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.Indent = 3;
            xls.SetCellFormat(158, 5, xls.AddFormat(fmt));
            xls.SetCellValue(158, 5, "Borbon");

            fmt = xls.GetCellVisibleFormatDef(158, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(158, 6, xls.AddFormat(fmt));
            xls.SetCellValue(158, 6, 0.4);

            fmt = xls.GetCellVisibleFormatDef(158, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(158, 7, xls.AddFormat(fmt));
            xls.SetCellValue(158, 7, "INPUTS\n%");

            fmt = xls.GetCellVisibleFormatDef(158, 8);
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(158, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(158, 13);
            fmt.Format = "0.00";
            xls.SetCellFormat(158, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(159, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(159, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(159, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.HAlignment = THFlxAlignment.left;
            xls.SetCellFormat(159, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(159, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.Indent = 3;
            xls.SetCellFormat(159, 5, xls.AddFormat(fmt));
            xls.SetCellValue(159, 5, "Catimore");

            fmt = xls.GetCellVisibleFormatDef(159, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(159, 6, xls.AddFormat(fmt));
            xls.SetCellValue(159, 6, 0.003);

            fmt = xls.GetCellVisibleFormatDef(159, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(159, 7, xls.AddFormat(fmt));
            xls.SetCellValue(159, 7, "INPUTS\n%");

            fmt = xls.GetCellVisibleFormatDef(159, 8);
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(159, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(159, 13);
            fmt.Format = "0.00";
            xls.SetCellFormat(159, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(159, 14);
            fmt.Format = "0.00";
            xls.SetCellFormat(159, 14, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(159, 15);
            fmt.Format = "0.00";
            xls.SetCellFormat(159, 15, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(160, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(160, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(160, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.HAlignment = THFlxAlignment.left;
            xls.SetCellFormat(160, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(160, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.Indent = 3;
            xls.SetCellFormat(160, 5, xls.AddFormat(fmt));
            xls.SetCellValue(160, 5, "Catuai");

            fmt = xls.GetCellVisibleFormatDef(160, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(160, 6, xls.AddFormat(fmt));
            xls.SetCellValue(160, 6, 0.002);

            fmt = xls.GetCellVisibleFormatDef(160, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(160, 7, xls.AddFormat(fmt));
            xls.SetCellValue(160, 7, "INPUTS\n%");

            fmt = xls.GetCellVisibleFormatDef(160, 8);
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(160, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(160, 13);
            fmt.Format = "0.00";
            xls.SetCellFormat(160, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(161, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(161, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(161, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.HAlignment = THFlxAlignment.left;
            xls.SetCellFormat(161, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(161, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.Indent = 3;
            xls.SetCellFormat(161, 5, xls.AddFormat(fmt));
            xls.SetCellValue(161, 5, "Caturra");

            fmt = xls.GetCellVisibleFormatDef(161, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(161, 6, xls.AddFormat(fmt));
            xls.SetCellValue(161, 6, 0.053);

            fmt = xls.GetCellVisibleFormatDef(161, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(161, 7, xls.AddFormat(fmt));
            xls.SetCellValue(161, 7, "INPUTS\n%");

            fmt = xls.GetCellVisibleFormatDef(161, 8);
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(161, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(161, 13);
            fmt.Format = "0.00";
            xls.SetCellFormat(161, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(161, 15);
            fmt.Format = "0.00";
            xls.SetCellFormat(161, 15, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(162, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(162, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(162, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.HAlignment = THFlxAlignment.left;
            xls.SetCellFormat(162, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(162, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.Indent = 3;
            xls.SetCellFormat(162, 5, xls.AddFormat(fmt));
            xls.SetCellValue(162, 5, "Colombia");

            fmt = xls.GetCellVisibleFormatDef(162, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(162, 6, xls.AddFormat(fmt));
            xls.SetCellValue(162, 6, 0);

            fmt = xls.GetCellVisibleFormatDef(162, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(162, 7, xls.AddFormat(fmt));
            xls.SetCellValue(162, 7, "INPUTS\n%");

            fmt = xls.GetCellVisibleFormatDef(162, 8);
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(162, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(162, 13);
            fmt.Format = "0.00";
            xls.SetCellFormat(162, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(162, 14);
            fmt.Format = "0.00";
            xls.SetCellFormat(162, 14, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(163, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(163, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(163, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.HAlignment = THFlxAlignment.left;
            xls.SetCellFormat(163, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(163, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.Indent = 3;
            xls.SetCellFormat(163, 5, xls.AddFormat(fmt));
            xls.SetCellValue(163, 5, "Costa Rica");

            fmt = xls.GetCellVisibleFormatDef(163, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(163, 6, xls.AddFormat(fmt));
            xls.SetCellValue(163, 6, 0);

            fmt = xls.GetCellVisibleFormatDef(163, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(163, 7, xls.AddFormat(fmt));
            xls.SetCellValue(163, 7, "INPUTS\n%");

            fmt = xls.GetCellVisibleFormatDef(163, 8);
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(163, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(163, 13);
            fmt.Format = "0.00";
            xls.SetCellFormat(163, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(164, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(164, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(164, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.HAlignment = THFlxAlignment.left;
            xls.SetCellFormat(164, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(164, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.Indent = 3;
            xls.SetCellFormat(164, 5, xls.AddFormat(fmt));
            xls.SetCellValue(164, 5, "Castillo");

            fmt = xls.GetCellVisibleFormatDef(164, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(164, 6, xls.AddFormat(fmt));
            xls.SetCellValue(164, 6, 0.003);

            fmt = xls.GetCellVisibleFormatDef(164, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(164, 7, xls.AddFormat(fmt));
            xls.SetCellValue(164, 7, "INPUTS\n%");

            fmt = xls.GetCellVisibleFormatDef(164, 8);
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(164, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(164, 13);
            fmt.Format = "0.00";
            xls.SetCellFormat(164, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(164, 14);
            fmt.Format = "0.00";
            xls.SetCellFormat(164, 14, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(165, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(165, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(165, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.HAlignment = THFlxAlignment.left;
            xls.SetCellFormat(165, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(165, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.Indent = 3;
            xls.SetCellFormat(165, 5, xls.AddFormat(fmt));
            xls.SetCellValue(165, 5, "Giesha");

            fmt = xls.GetCellVisibleFormatDef(165, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(165, 6, xls.AddFormat(fmt));
            xls.SetCellValue(165, 6, 0);

            fmt = xls.GetCellVisibleFormatDef(165, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(165, 7, xls.AddFormat(fmt));
            xls.SetCellValue(165, 7, "INPUTS\n%");

            fmt = xls.GetCellVisibleFormatDef(165, 8);
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(165, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(165, 13);
            fmt.Format = "0.00";
            xls.SetCellFormat(165, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(166, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(166, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(166, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.HAlignment = THFlxAlignment.left;
            xls.SetCellFormat(166, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(166, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.Indent = 3;
            xls.SetCellFormat(166, 5, xls.AddFormat(fmt));
            xls.SetCellValue(166, 5, "Icafe 90");

            fmt = xls.GetCellVisibleFormatDef(166, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(166, 6, xls.AddFormat(fmt));
            xls.SetCellValue(166, 6, 0);

            fmt = xls.GetCellVisibleFormatDef(166, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(166, 7, xls.AddFormat(fmt));
            xls.SetCellValue(166, 7, "INPUTS\n%");

            fmt = xls.GetCellVisibleFormatDef(166, 8);
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(166, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(167, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(167, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(167, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.HAlignment = THFlxAlignment.left;
            xls.SetCellFormat(167, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(167, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.Indent = 3;
            xls.SetCellFormat(167, 5, xls.AddFormat(fmt));
            xls.SetCellValue(167, 5, "Icatu");

            fmt = xls.GetCellVisibleFormatDef(167, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(167, 6, xls.AddFormat(fmt));
            xls.SetCellValue(167, 6, 0);

            fmt = xls.GetCellVisibleFormatDef(167, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(167, 7, xls.AddFormat(fmt));
            xls.SetCellValue(167, 7, "INPUTS\n%");

            fmt = xls.GetCellVisibleFormatDef(167, 8);
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(167, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(168, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(168, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(168, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.HAlignment = THFlxAlignment.left;
            xls.SetCellFormat(168, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(168, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.Indent = 3;
            xls.SetCellFormat(168, 5, xls.AddFormat(fmt));
            xls.SetCellValue(168, 5, "Lempira");

            fmt = xls.GetCellVisibleFormatDef(168, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(168, 6, xls.AddFormat(fmt));
            xls.SetCellValue(168, 6, 0);

            fmt = xls.GetCellVisibleFormatDef(168, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(168, 7, xls.AddFormat(fmt));
            xls.SetCellValue(168, 7, "INPUTS\n%");

            fmt = xls.GetCellVisibleFormatDef(168, 8);
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(168, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(169, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(169, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(169, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.HAlignment = THFlxAlignment.left;
            xls.SetCellFormat(169, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(169, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.Indent = 3;
            xls.SetCellFormat(169, 5, xls.AddFormat(fmt));
            xls.SetCellValue(169, 5, "Maragogype o Marago");

            fmt = xls.GetCellVisibleFormatDef(169, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(169, 6, xls.AddFormat(fmt));
            xls.SetCellValue(169, 6, 0);

            fmt = xls.GetCellVisibleFormatDef(169, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(169, 7, xls.AddFormat(fmt));
            xls.SetCellValue(169, 7, "INPUTS\n%");

            fmt = xls.GetCellVisibleFormatDef(169, 8);
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(169, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(170, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(170, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(170, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.Indent = 3;
            xls.SetCellFormat(170, 5, xls.AddFormat(fmt));
            xls.SetCellValue(170, 5, "Pacamara");

            fmt = xls.GetCellVisibleFormatDef(170, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(170, 6, xls.AddFormat(fmt));
            xls.SetCellValue(170, 6, 0);

            fmt = xls.GetCellVisibleFormatDef(170, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(170, 7, xls.AddFormat(fmt));
            xls.SetCellValue(170, 7, "INPUTS\n%");

            fmt = xls.GetCellVisibleFormatDef(170, 8);
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(170, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(171, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(171, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(171, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.Indent = 3;
            xls.SetCellFormat(171, 5, xls.AddFormat(fmt));
            xls.SetCellValue(171, 5, "Pache");

            fmt = xls.GetCellVisibleFormatDef(171, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(171, 6, xls.AddFormat(fmt));
            xls.SetCellValue(171, 6, 0);

            fmt = xls.GetCellVisibleFormatDef(171, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(171, 7, xls.AddFormat(fmt));
            xls.SetCellValue(171, 7, "INPUTS\n%");

            fmt = xls.GetCellVisibleFormatDef(171, 8);
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(171, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(172, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(172, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(172, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.Indent = 3;
            xls.SetCellFormat(172, 5, xls.AddFormat(fmt));
            xls.SetCellValue(172, 5, "Parainema");

            fmt = xls.GetCellVisibleFormatDef(172, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(172, 6, xls.AddFormat(fmt));
            xls.SetCellValue(172, 6, 0);

            fmt = xls.GetCellVisibleFormatDef(172, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(172, 7, xls.AddFormat(fmt));
            xls.SetCellValue(172, 7, "INPUTS\n%");

            fmt = xls.GetCellVisibleFormatDef(172, 8);
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(172, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(173, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(173, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(173, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.Indent = 3;
            xls.SetCellFormat(173, 5, xls.AddFormat(fmt));
            xls.SetCellValue(173, 5, "Suprema");

            fmt = xls.GetCellVisibleFormatDef(173, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(173, 6, xls.AddFormat(fmt));
            xls.SetCellValue(173, 6, 0);

            fmt = xls.GetCellVisibleFormatDef(173, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(173, 7, xls.AddFormat(fmt));
            xls.SetCellValue(173, 7, "INPUTS\n%");

            fmt = xls.GetCellVisibleFormatDef(173, 8);
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(173, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(174, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(174, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(174, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.Indent = 3;
            xls.SetCellFormat(174, 5, xls.AddFormat(fmt));
            xls.SetCellValue(174, 5, "Tipico");

            fmt = xls.GetCellVisibleFormatDef(174, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(174, 6, xls.AddFormat(fmt));
            xls.SetCellValue(174, 6, 0);

            fmt = xls.GetCellVisibleFormatDef(174, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(174, 7, xls.AddFormat(fmt));
            xls.SetCellValue(174, 7, "INPUTS\n%");

            fmt = xls.GetCellVisibleFormatDef(174, 8);
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(174, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(175, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(175, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(175, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.Indent = 3;
            xls.SetCellFormat(175, 5, xls.AddFormat(fmt));
            xls.SetCellValue(175, 5, "Villaserechi");

            fmt = xls.GetCellVisibleFormatDef(175, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(175, 6, xls.AddFormat(fmt));
            xls.SetCellValue(175, 6, 0);

            fmt = xls.GetCellVisibleFormatDef(175, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(175, 7, xls.AddFormat(fmt));
            xls.SetCellValue(175, 7, "INPUTS\n%");

            fmt = xls.GetCellVisibleFormatDef(175, 8);
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(175, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(176, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(176, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(176, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.Indent = 3;
            xls.SetCellFormat(176, 5, xls.AddFormat(fmt));
            xls.SetCellValue(176, 5, "Otra variedad:");

            fmt = xls.GetCellVisibleFormatDef(176, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(176, 6, xls.AddFormat(fmt));
            xls.SetCellValue(176, 6, 0.08);

            fmt = xls.GetCellVisibleFormatDef(176, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(176, 7, xls.AddFormat(fmt));
            xls.SetCellValue(176, 7, "INPUTS\n%");

            fmt = xls.GetCellVisibleFormatDef(176, 8);
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(176, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(177, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(177, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(177, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Style = TFlxFontStyles.Bold | TFlxFontStyles.Italic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(177, 5, xls.AddFormat(fmt));
            xls.SetCellValue(177, 5, "Metodos de Producción");

            fmt = xls.GetCellVisibleFormatDef(177, 7);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            xls.SetCellFormat(177, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(178, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(178, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(178, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(178, 5, xls.AddFormat(fmt));
            xls.SetCellValue(178, 5, "Finca química");

            fmt = xls.GetCellVisibleFormatDef(178, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(178, 6, xls.AddFormat(fmt));
            xls.SetCellValue(178, 6, new TFormula("='Inputs 1.0_metric_currency'!D10"));

            fmt = xls.GetCellVisibleFormatDef(178, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(178, 7, xls.AddFormat(fmt));
            xls.SetCellValue(178, 7, "INPUTS\n(yes, no)");

            fmt = xls.GetCellVisibleFormatDef(179, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(179, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(179, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(179, 5, xls.AddFormat(fmt));
            xls.SetCellValue(179, 5, "Finca orgánica");

            fmt = xls.GetCellVisibleFormatDef(179, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(179, 6, xls.AddFormat(fmt));
            xls.SetCellValue(179, 6, new TFormula("='Inputs 1.0_metric_currency'!D11"));

            fmt = xls.GetCellVisibleFormatDef(179, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(179, 7, xls.AddFormat(fmt));
            xls.SetCellValue(179, 7, "INPUTS\n(yes, no)");

            fmt = xls.GetCellVisibleFormatDef(180, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(180, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(180, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(180, 5, xls.AddFormat(fmt));
            xls.SetCellValue(180, 5, "Transición");

            fmt = xls.GetCellVisibleFormatDef(180, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(180, 6, xls.AddFormat(fmt));
            xls.SetCellValue(180, 6, new TFormula("='Inputs 1.0_metric_currency'!D12"));

            fmt = xls.GetCellVisibleFormatDef(180, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(180, 7, xls.AddFormat(fmt));
            xls.SetCellValue(180, 7, "INPUTS\n(yes, no)");

            fmt = xls.GetCellVisibleFormatDef(181, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(181, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(181, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Style = TFlxFontStyles.Bold | TFlxFontStyles.Italic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(181, 5, xls.AddFormat(fmt));
            xls.SetCellValue(181, 5, "Tipo de café producido y llevado a la asociación (Porcentaje):");

            fmt = xls.GetCellVisibleFormatDef(181, 7);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            xls.SetCellFormat(181, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(182, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(182, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(182, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.left;
            xls.SetCellFormat(182, 5, xls.AddFormat(fmt));
            xls.SetCellValue(182, 5, "Cereza");

            fmt = xls.GetCellVisibleFormatDef(182, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(182, 6, xls.AddFormat(fmt));
            xls.SetCellValue(182, 6, 100);

            fmt = xls.GetCellVisibleFormatDef(182, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(182, 7, xls.AddFormat(fmt));
            xls.SetCellValue(182, 7, "INPUTS\n%");

            fmt = xls.GetCellVisibleFormatDef(182, 8);
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(182, 8, xls.AddFormat(fmt));
            xls.SetCellValue(182, 8, "Sum = 100");

            fmt = xls.GetCellVisibleFormatDef(183, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(183, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(183, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.left;
            xls.SetCellFormat(183, 5, xls.AddFormat(fmt));
            xls.SetCellValue(183, 5, "Pergamino húmedo");

            fmt = xls.GetCellVisibleFormatDef(183, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(183, 6, xls.AddFormat(fmt));
            xls.SetCellValue(183, 6, 0);

            fmt = xls.GetCellVisibleFormatDef(183, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(183, 7, xls.AddFormat(fmt));
            xls.SetCellValue(183, 7, "INPUTS\n%");

            fmt = xls.GetCellVisibleFormatDef(183, 8);
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(183, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(184, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(184, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(184, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.left;
            xls.SetCellFormat(184, 5, xls.AddFormat(fmt));
            xls.SetCellValue(184, 5, "Pergamino seco");

            fmt = xls.GetCellVisibleFormatDef(184, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(184, 6, xls.AddFormat(fmt));
            xls.SetCellValue(184, 6, 0);

            fmt = xls.GetCellVisibleFormatDef(184, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(184, 7, xls.AddFormat(fmt));
            xls.SetCellValue(184, 7, "INPUTS\n%");

            fmt = xls.GetCellVisibleFormatDef(184, 8);
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(184, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(185, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(185, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(185, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.left;
            xls.SetCellFormat(185, 5, xls.AddFormat(fmt));
            xls.SetCellValue(185, 5, "Trillado");

            fmt = xls.GetCellVisibleFormatDef(185, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(185, 6, xls.AddFormat(fmt));
            xls.SetCellValue(185, 6, 0);

            fmt = xls.GetCellVisibleFormatDef(185, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(185, 7, xls.AddFormat(fmt));
            xls.SetCellValue(185, 7, "INPUTS\n%");

            fmt = xls.GetCellVisibleFormatDef(185, 8);
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(185, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(186, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(186, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(186, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.WrapText = true;
            xls.SetCellFormat(186, 5, xls.AddFormat(fmt));
            xls.SetCellValue(186, 5, "¿Vende usted parte de su café a otro comprador distinto de la asociación?");
            xls.SetCellValue(186, 6, 1);

            fmt = xls.GetCellVisibleFormatDef(186, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(186, 7, xls.AddFormat(fmt));
            xls.SetCellValue(186, 7, "INPUTS\n(yes = 1, no = 0)");

            fmt = xls.GetCellVisibleFormatDef(187, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(187, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(187, 4);
            fmt.WrapText = true;
            xls.SetCellFormat(187, 4, xls.AddFormat(fmt));
            xls.SetCellValue(187, 4, "* misma que la anterior");

            fmt = xls.GetCellVisibleFormatDef(187, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.WrapText = true;
            xls.SetCellFormat(187, 5, xls.AddFormat(fmt));
            xls.SetCellValue(187, 5, "¿Vende usted parte de su café a otro comprador distinto de la asociación?");

            fmt = xls.GetCellVisibleFormatDef(187, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(187, 6, xls.AddFormat(fmt));
            xls.SetCellValue(187, 6, 100);

            fmt = xls.GetCellVisibleFormatDef(187, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(187, 7, xls.AddFormat(fmt));
            xls.SetCellValue(187, 7, "INPUTS\n%");

            fmt = xls.GetCellVisibleFormatDef(188, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(188, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(188, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Style = TFlxFontStyles.Bold | TFlxFontStyles.Italic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(188, 5, xls.AddFormat(fmt));
            xls.SetCellValue(188, 5, "Actividades en la parcela");

            fmt = xls.GetCellVisibleFormatDef(188, 7);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            xls.SetCellFormat(188, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(189, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(189, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(189, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(189, 5, xls.AddFormat(fmt));
            xls.SetCellValue(189, 5, "¿Construye  el germinador o germinador? (1=sí, 0=no))");
            xls.SetCellValue(189, 6, 1);

            fmt = xls.GetCellVisibleFormatDef(189, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(189, 7, xls.AddFormat(fmt));
            xls.SetCellValue(189, 7, "INPUTS\n(yes = 1, no = 0)");

            fmt = xls.GetCellVisibleFormatDef(190, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(190, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(190, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.Indent = 2;
            xls.SetCellFormat(190, 5, xls.AddFormat(fmt));
            xls.SetCellValue(190, 5, "Valor estimado germinador");
            xls.SetCellValue(190, 6, 1192.4);

            fmt = xls.GetCellVisibleFormatDef(190, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(190, 7, xls.AddFormat(fmt));
            xls.SetCellValue(190, 7, "INPUTS\n#");

            fmt = xls.GetCellVisibleFormatDef(191, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(191, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(191, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(191, 5, xls.AddFormat(fmt));
            xls.SetCellValue(191, 5, "¿Construye  el vivero? (1=sí, 0=no))");
            xls.SetCellValue(191, 6, 1);

            fmt = xls.GetCellVisibleFormatDef(191, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(191, 7, xls.AddFormat(fmt));
            xls.SetCellValue(191, 7, "INPUTS\n(yes = 1, no = 0)");

            fmt = xls.GetCellVisibleFormatDef(192, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(192, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(192, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.Indent = 2;
            xls.SetCellFormat(192, 5, xls.AddFormat(fmt));
            xls.SetCellValue(192, 5, "Valor estimado vivero (de la estructura)");
            xls.SetCellValue(192, 6, 6268);

            fmt = xls.GetCellVisibleFormatDef(192, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(192, 7, xls.AddFormat(fmt));
            xls.SetCellValue(192, 7, "INPUTS\n#");

            fmt = xls.GetCellVisibleFormatDef(193, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(193, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(193, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(193, 5, xls.AddFormat(fmt));
            xls.SetCellValue(193, 5, "¿Compra  las plantulas o plantones? (1=sí, 0=no)");
            xls.SetCellValue(193, 6, 1);

            fmt = xls.GetCellVisibleFormatDef(193, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(193, 7, xls.AddFormat(fmt));
            xls.SetCellValue(193, 7, "INPUTS\n(yes = 1, no = 0)");

            fmt = xls.GetCellVisibleFormatDef(194, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(194, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(194, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.Indent = 2;
            xls.SetCellFormat(194, 5, xls.AddFormat(fmt));

            Runs = new TRTFRun[1];
            Runs[0].FirstChar = 19;
            fnt = xls.GetDefaultFont;
            fnt.Name = "Arial";
            fnt.Size20 = 280;
            fnt.Color = TExcelColor.Automatic;
            fnt.Underline = TFlxUnderline.Single;
            fnt.Scheme = TFontScheme.None;
            Runs[0].FontIndex = xls.AddFont(fnt);
            xls.SetCellValue(194, 5, new TRichString("Valor estimado  por planta", Runs, xls));
            //We could also have used: xls.SetCellFromHtml(194, 5, "Valor estimado &nbsp;por<u>&nbsp;planta</u>")

            xls.SetCellValue(194, 6, 5.17);

            fmt = xls.GetCellVisibleFormatDef(194, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(194, 7, xls.AddFormat(fmt));
            xls.SetCellValue(194, 7, "INPUTS\n#");

            fmt = xls.GetCellVisibleFormatDef(195, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(195, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(195, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Style = TFlxFontStyles.Bold | TFlxFontStyles.Italic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(195, 5, xls.AddFormat(fmt));
            xls.SetCellValue(195, 5, "Ingresos");

            fmt = xls.GetCellVisibleFormatDef(195, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 360;
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            xls.SetCellFormat(195, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(196, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(196, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(196, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.WrapText = true;
            xls.SetCellFormat(196, 5, xls.AddFormat(fmt));

            Runs = new TRTFRun[2];
            Runs[0].FirstChar = 43;
            fnt = xls.GetDefaultFont;
            fnt.Name = "Arial";
            fnt.Size20 = 280;
            fnt.Style = TFlxFontStyles.Bold;
            fnt.Scheme = TFontScheme.None;
            Runs[0].FontIndex = xls.AddFont(fnt);
            Runs[1].FirstChar = 62;
            fnt = xls.GetDefaultFont;
            fnt.Name = "Arial";
            fnt.Size20 = 280;
            fnt.Scheme = TFontScheme.None;
            Runs[1].FontIndex = xls.AddFont(fnt);
            xls.SetCellValue(196, 5, new TRichString("Cuál fue el precio promedio por QUINTAL de café pergamino seco, que usted recibió"
            + " en la última cosecha? ", Runs, xls));
            //We could also have used: xls.SetCellFromHtml(196, 5, "Cu&aacute;l fue el precio promedio por QUINTAL de&nbsp;<b>caf&eacute; pergamino seco</b>,"
            //+" que usted recibi&oacute; en la &uacute;ltima cosecha?&nbsp;")


    fmt = xls.GetCellVisibleFormatDef(196, 6);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent2, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(196, 6, xls.AddFormat(fmt));
            xls.SetCellValue(196, 6, new TFormula("='Inputs 1.0_metric_currency'!$D$19"));

            fmt = xls.GetCellVisibleFormatDef(196, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(196, 7, xls.AddFormat(fmt));
            xls.SetCellValue(196, 7, "INPUTS\n#");

            fmt = xls.GetCellVisibleFormatDef(196, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(196, 8, xls.AddFormat(fmt));
            xls.SetCellValue(196, 8, "no se usa para calculos");

            fmt = xls.GetCellVisibleFormatDef(197, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(197, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(197, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.WrapText = true;
            xls.SetCellFormat(197, 5, xls.AddFormat(fmt));
            xls.SetCellValue(197, 5, "Ha recibido en algún momento algún premio asociado a su producción de café? (Ejemplo:"
            + " Comercio Justo, orgánico, premio de taza, factor de rendimiento). ");
            xls.SetCellValue(197, 6, 1);

            fmt = xls.GetCellVisibleFormatDef(197, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(197, 7, xls.AddFormat(fmt));
            xls.SetCellValue(197, 7, "INPUTS\n(yes = 1, no = 0)");

            fmt = xls.GetCellVisibleFormatDef(197, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(197, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(198, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(198, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(198, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.WrapText = true;
            xls.SetCellFormat(198, 5, xls.AddFormat(fmt));
            xls.SetCellValue(198, 5, "¿Cuál fue el valor por KILO del premio Comercio Justo ");
            xls.SetCellValue(198, 6, 164.4);

            fmt = xls.GetCellVisibleFormatDef(198, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(198, 7, xls.AddFormat(fmt));
            xls.SetCellValue(198, 7, "INPUTS\n#");

            fmt = xls.GetCellVisibleFormatDef(198, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(198, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(199, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(199, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(199, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.WrapText = true;
            xls.SetCellFormat(199, 5, xls.AddFormat(fmt));
            xls.SetCellValue(199, 5, "¿Cuál fue el valor por KILO del premio Orgánico");
            xls.SetCellValue(199, 6, 0);

            fmt = xls.GetCellVisibleFormatDef(199, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(199, 7, xls.AddFormat(fmt));
            xls.SetCellValue(199, 7, "INPUTS\n#");

            fmt = xls.GetCellVisibleFormatDef(199, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(199, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(200, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(200, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(200, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.WrapText = true;
            xls.SetCellFormat(200, 5, xls.AddFormat(fmt));
            xls.SetCellValue(200, 5, "¿Cuál fue el valor por KILO del premio Cooperativa");
            xls.SetCellValue(200, 6, 205);

            fmt = xls.GetCellVisibleFormatDef(200, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(200, 7, xls.AddFormat(fmt));
            xls.SetCellValue(200, 7, "INPUTS\n#");

            fmt = xls.GetCellVisibleFormatDef(200, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(200, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(201, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(201, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(201, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.WrapText = true;
            xls.SetCellFormat(201, 5, xls.AddFormat(fmt));
            xls.SetCellValue(201, 5, "¿Cuál fue el valor por KILO de otros premios?");
            xls.SetCellValue(201, 6, 0);

            fmt = xls.GetCellVisibleFormatDef(201, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(201, 7, xls.AddFormat(fmt));
            xls.SetCellValue(201, 7, "INPUTS\n#");

            fmt = xls.GetCellVisibleFormatDef(201, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(201, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(202, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(202, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(202, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.WrapText = true;
            xls.SetCellFormat(202, 5, xls.AddFormat(fmt));
            xls.SetCellValue(202, 5, "Si va a contratar a una persona por días, ¿cuál es el valor promedio que le pagaría"
            + " por día? (Ej: Jornal, tarea que dura un día)");
            xls.SetCellValue(202, 6, new TFormula("='Inputs 1.0_metric_currency'!D14"));

            fmt = xls.GetCellVisibleFormatDef(202, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(202, 7, xls.AddFormat(fmt));
            xls.SetCellValue(202, 7, "INPUTS\n#");

            fmt = xls.GetCellVisibleFormatDef(202, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(202, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(203, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(203, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(203, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.WrapText = true;
            xls.SetCellFormat(203, 5, xls.AddFormat(fmt));
            xls.SetCellValue(203, 5, "Si fuera a ser contratado por alguién hoy, ¿cuál sería el valor aproximado que a "
            + " le pagarían?");
            xls.SetCellValue(203, 6, 107);

            fmt = xls.GetCellVisibleFormatDef(203, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(203, 7, xls.AddFormat(fmt));
            xls.SetCellValue(203, 7, "INPUTS\n#");

            fmt = xls.GetCellVisibleFormatDef(203, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(203, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(204, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(204, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(204, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.WrapText = true;
            xls.SetCellFormat(204, 5, xls.AddFormat(fmt));
            xls.SetCellValue(204, 5, "Alimenta usted a sus trabajadores aparte de pagarles el jornal? ¿Cuál es el valor"
            + " estimado?");
            xls.SetCellValue(204, 6, 0);

            fmt = xls.GetCellVisibleFormatDef(204, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(204, 7, xls.AddFormat(fmt));
            xls.SetCellValue(204, 7, "INPUTS\n#");

            fmt = xls.GetCellVisibleFormatDef(204, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(204, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(205, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(205, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(205, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.WrapText = true;
            xls.SetCellFormat(205, 5, xls.AddFormat(fmt));
            xls.SetCellValue(205, 5, "¿Cuál es el salario mínimo mensual vigente actualmente?");
            xls.SetCellValue(205, 6, 817);

            fmt = xls.GetCellVisibleFormatDef(205, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(205, 7, xls.AddFormat(fmt));
            xls.SetCellValue(205, 7, "INPUTS\n#");

            fmt = xls.GetCellVisibleFormatDef(205, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(205, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(206, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(206, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(206, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.WrapText = true;
            xls.SetCellFormat(206, 5, xls.AddFormat(fmt));
            xls.SetCellValue(206, 5, "¿Últimamente se ha visto su finca fuertmente afectada por plagas, enfermedades o desastres"
            + " naturales en algún año en particular? ¿Cuál año?");
            xls.SetCellValue(206, 6, 2012);

            fmt = xls.GetCellVisibleFormatDef(206, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(206, 7, xls.AddFormat(fmt));
            xls.SetCellValue(206, 7, "INPUTS\n#");

            fmt = xls.GetCellVisibleFormatDef(206, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(206, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(207, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(207, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(207, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.WrapText = true;
            xls.SetCellFormat(207, 5, xls.AddFormat(fmt));
            xls.SetCellValue(207, 5, "¿En qué porcentaje se redujo su producción como consecuencia de este evento particular?");
            xls.SetCellValue(207, 6, 0.5);

            fmt = xls.GetCellVisibleFormatDef(207, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(207, 7, xls.AddFormat(fmt));
            xls.SetCellValue(207, 7, "INPUTS\n%");

            fmt = xls.GetCellVisibleFormatDef(207, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(207, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(208, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(208, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(208, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Style = TFlxFontStyles.Italic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.WrapText = true;
            xls.SetCellFormat(208, 5, xls.AddFormat(fmt));
            xls.SetCellValue(208, 5, "¿Que alternativas utilizó  para sobrepasar el choque en los ingresos que ese evento"
            + " representó? (Ej: crédito, venta de lote)");

            fmt = xls.GetCellVisibleFormatDef(208, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(208, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(208, 7);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            xls.SetCellFormat(208, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(208, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(208, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(209, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(209, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(209, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.right;
            fmt.WrapText = true;
            xls.SetCellFormat(209, 5, xls.AddFormat(fmt));
            xls.SetCellValue(209, 5, "Préstamos");

            fmt = xls.GetCellVisibleFormatDef(209, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "0";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(209, 6, xls.AddFormat(fmt));
            xls.SetCellValue(209, 6, 1);

            fmt = xls.GetCellVisibleFormatDef(209, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(209, 7, xls.AddFormat(fmt));
            xls.SetCellValue(209, 7, "INPUTS\n(yes = 1, no = 0)");

            fmt = xls.GetCellVisibleFormatDef(209, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(209, 8, xls.AddFormat(fmt));
            xls.SetCellValue(209, 8, "no se usa para calculos");
            xls.SetCellValue(209, 9, new TFormula("=Budget_Supuestos!D432"));

            fmt = xls.GetCellVisibleFormatDef(210, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(210, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(210, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.right;
            fmt.WrapText = true;
            xls.SetCellFormat(210, 5, xls.AddFormat(fmt));
            xls.SetCellValue(210, 5, "Venta de activos (Lotes, tierra, acciones en la asociación)");

            fmt = xls.GetCellVisibleFormatDef(210, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "0";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(210, 6, xls.AddFormat(fmt));
            xls.SetCellValue(210, 6, 0);

            fmt = xls.GetCellVisibleFormatDef(210, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(210, 7, xls.AddFormat(fmt));
            xls.SetCellValue(210, 7, "INPUTS\n(yes = 1, no = 0)");

            fmt = xls.GetCellVisibleFormatDef(210, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(210, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(211, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(211, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(211, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.right;
            fmt.VAlignment = TVFlxAlignment.top;
            fmt.WrapText = true;
            xls.SetCellFormat(211, 5, xls.AddFormat(fmt));
            xls.SetCellValue(211, 5, "Trabajo particular");

            fmt = xls.GetCellVisibleFormatDef(211, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "0";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(211, 6, xls.AddFormat(fmt));
            xls.SetCellValue(211, 6, 0);

            fmt = xls.GetCellVisibleFormatDef(211, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(211, 7, xls.AddFormat(fmt));
            xls.SetCellValue(211, 7, "INPUTS\n(yes = 1, no = 0)");

            fmt = xls.GetCellVisibleFormatDef(211, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(211, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(212, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(212, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(212, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(212, 5, xls.AddFormat(fmt));
            xls.SetCellValue(212, 5, "Uso de ahorros");

            fmt = xls.GetCellVisibleFormatDef(212, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "0";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(212, 6, xls.AddFormat(fmt));
            xls.SetCellValue(212, 6, 1);

            fmt = xls.GetCellVisibleFormatDef(212, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(212, 7, xls.AddFormat(fmt));
            xls.SetCellValue(212, 7, "INPUTS\n(yes = 1, no = 0)");

            fmt = xls.GetCellVisibleFormatDef(212, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(212, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(213, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(213, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(213, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.right;
            fmt.WrapText = true;
            xls.SetCellFormat(213, 5, xls.AddFormat(fmt));
            xls.SetCellValue(213, 5, "Renovación a otras variedades. ¿Cuál?");

            fmt = xls.GetCellVisibleFormatDef(213, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "0";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(213, 6, xls.AddFormat(fmt));
            xls.SetCellValue(213, 6, 0);

            fmt = xls.GetCellVisibleFormatDef(213, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(213, 7, xls.AddFormat(fmt));
            xls.SetCellValue(213, 7, "INPUTS\n(yes = 1, no = 0)");

            fmt = xls.GetCellVisibleFormatDef(213, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(213, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(214, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(214, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(214, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.right;
            fmt.WrapText = true;
            xls.SetCellFormat(214, 5, xls.AddFormat(fmt));
            xls.SetCellValue(214, 5, "Transición químico a orgánico");

            fmt = xls.GetCellVisibleFormatDef(214, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "0";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(214, 6, xls.AddFormat(fmt));
            xls.SetCellValue(214, 6, 0);

            fmt = xls.GetCellVisibleFormatDef(214, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(214, 7, xls.AddFormat(fmt));
            xls.SetCellValue(214, 7, "INPUTS\n(yes = 1, no = 0)");

            fmt = xls.GetCellVisibleFormatDef(214, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(214, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(215, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(215, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(215, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.right;
            fmt.WrapText = true;
            xls.SetCellFormat(215, 5, xls.AddFormat(fmt));
            xls.SetCellValue(215, 5, "Transición orgánico a químico");

            fmt = xls.GetCellVisibleFormatDef(215, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "0";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(215, 6, xls.AddFormat(fmt));
            xls.SetCellValue(215, 6, 0);

            fmt = xls.GetCellVisibleFormatDef(215, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(215, 7, xls.AddFormat(fmt));
            xls.SetCellValue(215, 7, "INPUTS\n(yes = 1, no = 0)");

            fmt = xls.GetCellVisibleFormatDef(215, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(215, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(216, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(216, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(216, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(216, 5, xls.AddFormat(fmt));
            xls.SetCellValue(216, 5, "Paso a otro cultivo");

            fmt = xls.GetCellVisibleFormatDef(216, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "0";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(216, 6, xls.AddFormat(fmt));
            xls.SetCellValue(216, 6, 0);

            fmt = xls.GetCellVisibleFormatDef(216, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(216, 7, xls.AddFormat(fmt));
            xls.SetCellValue(216, 7, "INPUTS\n(yes = 1, no = 0)");

            fmt = xls.GetCellVisibleFormatDef(216, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(216, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(217, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(217, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(217, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(217, 5, xls.AddFormat(fmt));
            xls.SetCellValue(217, 5, "Otros:");

            fmt = xls.GetCellVisibleFormatDef(217, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "0";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(217, 6, xls.AddFormat(fmt));
            xls.SetCellValue(217, 6, 0);

            fmt = xls.GetCellVisibleFormatDef(217, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(217, 7, xls.AddFormat(fmt));
            xls.SetCellValue(217, 7, "INPUTS\n(yes = 1, no = 0)");

            fmt = xls.GetCellVisibleFormatDef(217, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(217, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(218, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(218, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(218, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 360;
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(218, 5, xls.AddFormat(fmt));
            xls.SetCellValue(218, 5, "Ingresos indirectos");

            fmt = xls.GetCellVisibleFormatDef(218, 7);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            xls.SetCellFormat(218, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(219, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(219, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(219, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.WrapText = true;
            xls.SetCellFormat(219, 5, xls.AddFormat(fmt));
            xls.SetCellValue(219, 5, "¿Recibió otros ingresos por parte de la asociación DIFERENTES a préstamos? (Ej: Becas,"
            + " abonos, fertilizantes)");

            fmt = xls.GetCellVisibleFormatDef(219, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "0";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(219, 6, xls.AddFormat(fmt));
            xls.SetCellValue(219, 6, 0);

            fmt = xls.GetCellVisibleFormatDef(219, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(219, 7, xls.AddFormat(fmt));
            xls.SetCellValue(219, 7, "INPUTS\n(yes = 1, no = 0)");

            fmt = xls.GetCellVisibleFormatDef(220, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(220, 3, xls.AddFormat(fmt));
            xls.SetCellValue(220, 4, "sólo espacio para una cosa");

            fmt = xls.GetCellVisibleFormatDef(220, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(220, 5, xls.AddFormat(fmt));
            xls.SetCellValue(220, 5, "Descripción");
            xls.SetCellValue(220, 6, "extintor");

            fmt = xls.GetCellVisibleFormatDef(220, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(220, 7, xls.AddFormat(fmt));
            xls.SetCellValue(220, 7, "INPUTS\nAlphanumeric");

            fmt = xls.GetCellVisibleFormatDef(221, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(221, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(221, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(221, 5, xls.AddFormat(fmt));
            xls.SetCellValue(221, 5, "Por Cuántos años ha recibido esta ayuda");

            fmt = xls.GetCellVisibleFormatDef(221, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "0";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(221, 6, xls.AddFormat(fmt));
            xls.SetCellValue(221, 6, 4);

            fmt = xls.GetCellVisibleFormatDef(221, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(221, 7, xls.AddFormat(fmt));
            xls.SetCellValue(221, 7, "INPUTS\n#");

            fmt = xls.GetCellVisibleFormatDef(222, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(222, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(222, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(222, 5, xls.AddFormat(fmt));
            xls.SetCellValue(222, 5, "Valor anual aproximado de la ayuda que recbio");
            xls.SetCellValue(222, 6, 8255);

            fmt = xls.GetCellVisibleFormatDef(222, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(222, 7, xls.AddFormat(fmt));
            xls.SetCellValue(222, 7, "INPUTS\n#");

            fmt = xls.GetCellVisibleFormatDef(223, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(223, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(223, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(223, 5, xls.AddFormat(fmt));
            xls.SetCellValue(223, 5, "¿Recibió usted por parte de la asociación algún tipo de capacitación en los últimos"
            + " dos años?");

            fmt = xls.GetCellVisibleFormatDef(223, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "0";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(223, 6, xls.AddFormat(fmt));
            xls.SetCellValue(223, 6, 1);

            fmt = xls.GetCellVisibleFormatDef(223, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(223, 7, xls.AddFormat(fmt));
            xls.SetCellValue(223, 7, "INPUTS\n(yes = 1, no = 0)");

            fmt = xls.GetCellVisibleFormatDef(224, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(224, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(224, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(224, 5, xls.AddFormat(fmt));
            xls.SetCellValue(224, 5, "Temas tratados en la capacitacion (tres principales temas)");
            xls.SetCellValue(224, 6, "foliación");

            fmt = xls.GetCellVisibleFormatDef(224, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(224, 7, xls.AddFormat(fmt));
            xls.SetCellValue(224, 7, "INPUTS\nAlphanumeric");

            fmt = xls.GetCellVisibleFormatDef(225, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(225, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(225, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(225, 5, xls.AddFormat(fmt));
            xls.SetCellValue(225, 5, "Año");

            fmt = xls.GetCellVisibleFormatDef(225, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "0";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(225, 6, xls.AddFormat(fmt));
            xls.SetCellValue(225, 6, 2015);

            fmt = xls.GetCellVisibleFormatDef(225, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(225, 7, xls.AddFormat(fmt));
            xls.SetCellValue(225, 7, "INPUTS\n#");

            fmt = xls.GetCellVisibleFormatDef(226, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(226, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(226, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(226, 5, xls.AddFormat(fmt));
            xls.SetCellValue(226, 5, "número de años");

            fmt = xls.GetCellVisibleFormatDef(226, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "0";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(226, 6, xls.AddFormat(fmt));
            xls.SetCellValue(226, 6, 2);

            fmt = xls.GetCellVisibleFormatDef(226, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(226, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(227, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(227, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(227, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(227, 5, xls.AddFormat(fmt));

            Runs = new TRTFRun[1];
            Runs[0].FirstChar = 10;
            fnt = xls.GetDefaultFont;
            fnt.Name = "Arial";
            fnt.Size20 = 280;
            fnt.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fnt.Scheme = TFontScheme.None;
            Runs[0].FontIndex = xls.AddFont(fnt);
            xls.SetCellValue(227, 5, new TRichString("Intensidad o duracion en días de cada capacitacion", Runs, xls));
            //We could also have used: xls.SetCellFromHtml(227, 5, "Intensidad<font color = 'blue'>&nbsp;o duracion en d&iacute;as de cada capacitacion</font>")

            xls.SetCellValue(227, 6, 2.7);

            fmt = xls.GetCellVisibleFormatDef(227, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(227, 7, xls.AddFormat(fmt));
            xls.SetCellValue(227, 7, "INPUTS\n#");

            fmt = xls.GetCellVisibleFormatDef(228, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(228, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(228, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(228, 5, xls.AddFormat(fmt));
            xls.SetCellValue(228, 5, "Veces  al año");

            fmt = xls.GetCellVisibleFormatDef(228, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "0";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(228, 6, xls.AddFormat(fmt));
            xls.SetCellValue(228, 6, 1);

            fmt = xls.GetCellVisibleFormatDef(228, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(228, 7, xls.AddFormat(fmt));
            xls.SetCellValue(228, 7, "INPUTS\n#");

            fmt = xls.GetCellVisibleFormatDef(229, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(229, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(229, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 360;
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(229, 5, xls.AddFormat(fmt));
            xls.SetCellValue(229, 5, "Crédito");

            fmt = xls.GetCellVisibleFormatDef(229, 7);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            xls.SetCellFormat(229, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(230, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(230, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(230, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "0";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(230, 5, xls.AddFormat(fmt));
            xls.SetCellValue(230, 5, "Recibio algún prestamo por parte de la asociación?");
            xls.SetCellValue(230, 6, 1);

            fmt = xls.GetCellVisibleFormatDef(230, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(230, 7, xls.AddFormat(fmt));
            xls.SetCellValue(230, 7, "INPUTS\n(yes = 1, no = 0)");

            fmt = xls.GetCellVisibleFormatDef(231, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(231, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(231, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(231, 5, xls.AddFormat(fmt));
            xls.SetCellValue(231, 5, "Propósito");

            fmt = xls.GetCellVisibleFormatDef(231, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(231, 6, xls.AddFormat(fmt));
            xls.SetCellValue(231, 6, "fert");

            fmt = xls.GetCellVisibleFormatDef(231, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(231, 7, xls.AddFormat(fmt));
            xls.SetCellValue(231, 7, "INPUTS\nAlphanumeric");

            fmt = xls.GetCellVisibleFormatDef(232, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(232, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(232, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(232, 5, xls.AddFormat(fmt));
            xls.SetCellValue(232, 5, "¿Cuándo recibió el préstamo? (Año / mes)");

            fmt = xls.GetCellVisibleFormatDef(232, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent4, 0.399975585192419);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(232, 6, xls.AddFormat(fmt));
            xls.SetCellValue(232, 6, 2015.4);

            fmt = xls.GetCellVisibleFormatDef(232, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(232, 7, xls.AddFormat(fmt));
            xls.SetCellValue(232, 7, "INPUTS\n#");

            fmt = xls.GetCellVisibleFormatDef(233, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(233, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(233, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(233, 5, xls.AddFormat(fmt));
            xls.SetCellValue(233, 5, "Monto en moneda local");

            fmt = xls.GetCellVisibleFormatDef(233, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(233, 6, xls.AddFormat(fmt));
            xls.SetCellValue(233, 6, 14000);

            fmt = xls.GetCellVisibleFormatDef(233, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(233, 7, xls.AddFormat(fmt));
            xls.SetCellValue(233, 7, "INPUTS\n#");

            fmt = xls.GetCellVisibleFormatDef(234, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(234, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(234, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(234, 5, xls.AddFormat(fmt));
            xls.SetCellValue(234, 5, "¿Cuándo termina o terminó de pagar el préstamo? (Año / mes)");

            fmt = xls.GetCellVisibleFormatDef(234, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent4, 0.399975585192419);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(234, 6, xls.AddFormat(fmt));
            xls.SetCellValue(234, 6, 2017);

            fmt = xls.GetCellVisibleFormatDef(234, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(234, 7, xls.AddFormat(fmt));
            xls.SetCellValue(234, 7, "INPUTS\n#");

            fmt = xls.GetCellVisibleFormatDef(235, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(235, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(235, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(235, 5, xls.AddFormat(fmt));
            xls.SetCellValue(235, 5, "Tiempo del préstamo");

            fmt = xls.GetCellVisibleFormatDef(235, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent4, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(235, 6, xls.AddFormat(fmt));
            xls.SetCellValue(235, 6, new TFormula("=F234-F232"));

            fmt = xls.GetCellVisibleFormatDef(235, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(235, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(236, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(236, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(236, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(236, 5, xls.AddFormat(fmt));
            xls.SetCellValue(236, 5, "Pagos del préstamo");
            xls.SetCellValue(236, 6, "mensual");

            fmt = xls.GetCellVisibleFormatDef(236, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(236, 7, xls.AddFormat(fmt));
            xls.SetCellValue(236, 7, "INPUTS\n(mensual, anual)");

            fmt = xls.GetCellVisibleFormatDef(237, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(237, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(237, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(237, 5, xls.AddFormat(fmt));
            xls.SetCellValue(237, 5, "Monto del pago");

            fmt = xls.GetCellVisibleFormatDef(237, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(237, 6, xls.AddFormat(fmt));
            xls.SetCellValue(237, 6, 300);

            fmt = xls.GetCellVisibleFormatDef(237, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(237, 7, xls.AddFormat(fmt));
            xls.SetCellValue(237, 7, "INPUTS\n#");

            fmt = xls.GetCellVisibleFormatDef(238, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(238, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(238, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(238, 5, xls.AddFormat(fmt));
            xls.SetCellValue(238, 5, "Tasa de interés");
            xls.SetCellValue(238, 6, "mensual");

            fmt = xls.GetCellVisibleFormatDef(238, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(238, 7, xls.AddFormat(fmt));
            xls.SetCellValue(238, 7, "INPUTS\n(mensual, anual)");

            fmt = xls.GetCellVisibleFormatDef(239, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(239, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(239, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(239, 5, xls.AddFormat(fmt));
            xls.SetCellValue(239, 5, "Tasa de interés");

            fmt = xls.GetCellVisibleFormatDef(239, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(239, 6, xls.AddFormat(fmt));
            xls.SetCellValue(239, 6, 1.01);

            fmt = xls.GetCellVisibleFormatDef(239, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(239, 7, xls.AddFormat(fmt));
            xls.SetCellValue(239, 7, "INPUTS\n#");

            fmt = xls.GetCellVisibleFormatDef(240, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(240, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(240, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "0";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(240, 5, xls.AddFormat(fmt));
            xls.SetCellValue(240, 5, "Recibio algún prestamo por parte de un Banco u otro prestamista? ");

            fmt = xls.GetCellVisibleFormatDef(240, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(240, 6, xls.AddFormat(fmt));
            xls.SetCellValue(240, 6, 1);

            fmt = xls.GetCellVisibleFormatDef(240, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(240, 7, xls.AddFormat(fmt));
            xls.SetCellValue(240, 7, "INPUTS\n(yes = 1, no = 0)");

            fmt = xls.GetCellVisibleFormatDef(241, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(241, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(241, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(241, 5, xls.AddFormat(fmt));
            xls.SetCellValue(241, 5, "Propósito");

            fmt = xls.GetCellVisibleFormatDef(241, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(241, 6, xls.AddFormat(fmt));
            xls.SetCellValue(241, 6, "inv");

            fmt = xls.GetCellVisibleFormatDef(241, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(241, 7, xls.AddFormat(fmt));
            xls.SetCellValue(241, 7, "INPUTS\nAlphanumeric");

            fmt = xls.GetCellVisibleFormatDef(242, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(242, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(242, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(242, 5, xls.AddFormat(fmt));
            xls.SetCellValue(242, 5, "¿Cuándo recibió el préstamo? (Año / mes)");

            fmt = xls.GetCellVisibleFormatDef(242, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent4, 0.399975585192419);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(242, 6, xls.AddFormat(fmt));
            xls.SetCellValue(242, 6, 2014.8);

            fmt = xls.GetCellVisibleFormatDef(242, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(242, 7, xls.AddFormat(fmt));
            xls.SetCellValue(242, 7, "INPUTS\n#");

            fmt = xls.GetCellVisibleFormatDef(243, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(243, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(243, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(243, 5, xls.AddFormat(fmt));
            xls.SetCellValue(243, 5, "Monto en moneda local");

            fmt = xls.GetCellVisibleFormatDef(243, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(243, 6, xls.AddFormat(fmt));
            xls.SetCellValue(243, 6, 5260);

            fmt = xls.GetCellVisibleFormatDef(243, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(243, 7, xls.AddFormat(fmt));
            xls.SetCellValue(243, 7, "INPUTS\n#");

            fmt = xls.GetCellVisibleFormatDef(244, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(244, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(244, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(244, 5, xls.AddFormat(fmt));
            xls.SetCellValue(244, 5, "¿Cuándo termina o terminó de pagar el préstamo? (Año / mes)");

            fmt = xls.GetCellVisibleFormatDef(244, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent4, 0.399975585192419);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(244, 6, xls.AddFormat(fmt));
            xls.SetCellValue(244, 6, 2017);

            fmt = xls.GetCellVisibleFormatDef(244, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(244, 7, xls.AddFormat(fmt));
            xls.SetCellValue(244, 7, "INPUTS\n#");

            fmt = xls.GetCellVisibleFormatDef(245, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(245, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(245, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(245, 5, xls.AddFormat(fmt));
            xls.SetCellValue(245, 5, "Tiempo del préstamo");

            fmt = xls.GetCellVisibleFormatDef(245, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent4, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(245, 6, xls.AddFormat(fmt));
            xls.SetCellValue(245, 6, new TFormula("=F244-F242"));

            fmt = xls.GetCellVisibleFormatDef(245, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(245, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(246, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(246, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(246, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(246, 5, xls.AddFormat(fmt));
            xls.SetCellValue(246, 5, "Pagos del préstamo");

            fmt = xls.GetCellVisibleFormatDef(246, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(246, 6, xls.AddFormat(fmt));
            xls.SetCellValue(246, 6, "mensual");

            fmt = xls.GetCellVisibleFormatDef(246, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(246, 7, xls.AddFormat(fmt));
            xls.SetCellValue(246, 7, "INPUTS\n(mensual, anual)");

            fmt = xls.GetCellVisibleFormatDef(247, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(247, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(247, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(247, 5, xls.AddFormat(fmt));
            xls.SetCellValue(247, 5, "Monto del pago");

            fmt = xls.GetCellVisibleFormatDef(247, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(247, 6, xls.AddFormat(fmt));
            xls.SetCellValue(247, 6, 230);

            fmt = xls.GetCellVisibleFormatDef(247, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(247, 7, xls.AddFormat(fmt));
            xls.SetCellValue(247, 7, "INPUTS\n#");

            fmt = xls.GetCellVisibleFormatDef(248, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(248, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(248, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(248, 5, xls.AddFormat(fmt));
            xls.SetCellValue(248, 5, "Tasa de interés");

            fmt = xls.GetCellVisibleFormatDef(248, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(248, 6, xls.AddFormat(fmt));
            xls.SetCellValue(248, 6, "mensual");

            fmt = xls.GetCellVisibleFormatDef(248, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(248, 7, xls.AddFormat(fmt));
            xls.SetCellValue(248, 7, "INPUTS\n(mensual, anual)");

            fmt = xls.GetCellVisibleFormatDef(249, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(249, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(249, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(249, 5, xls.AddFormat(fmt));
            xls.SetCellValue(249, 5, "Tasa de interés");

            fmt = xls.GetCellVisibleFormatDef(249, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(249, 6, xls.AddFormat(fmt));
            xls.SetCellValue(249, 6, 3.21);

            fmt = xls.GetCellVisibleFormatDef(249, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(249, 7, xls.AddFormat(fmt));
            xls.SetCellValue(249, 7, "INPUTS\n#");

            fmt = xls.GetCellVisibleFormatDef(250, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(250, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(250, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 400;
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(250, 5, xls.AddFormat(fmt));
            xls.SetCellValue(250, 5, "Costos materiales o insumos");

            fmt = xls.GetCellVisibleFormatDef(250, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 400;
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(250, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(250, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(250, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(250, 8);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 400;
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(250, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(250, 9);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 400;
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(250, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(250, 10);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 400;
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(250, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(251, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(251, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(251, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Style = TFlxFontStyles.Bold | TFlxFontStyles.Italic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(251, 5, xls.AddFormat(fmt));
            xls.SetCellValue(251, 5, "Materiales para el germinador (Año 0)");

            fmt = xls.GetCellVisibleFormatDef(251, 7);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            xls.SetCellFormat(251, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(252, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(252, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(252, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(252, 5, xls.AddFormat(fmt));
            xls.SetCellValue(252, 5, "Semilla");
            xls.SetCellValue(252, 6, 487);

            fmt = xls.GetCellVisibleFormatDef(252, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(252, 7, xls.AddFormat(fmt));
            xls.SetCellValue(252, 7, "INPUTS\n(required amount * cost per unit)");

            fmt = xls.GetCellVisibleFormatDef(253, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(253, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(253, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(253, 5, xls.AddFormat(fmt));
            xls.SetCellValue(253, 5, "Germinador/Marco semillero");
            xls.SetCellValue(253, 6, 430);

            fmt = xls.GetCellVisibleFormatDef(253, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(253, 7, xls.AddFormat(fmt));
            xls.SetCellValue(253, 7, "INPUTS\n(required amount * cost per unit)");

            fmt = xls.GetCellVisibleFormatDef(254, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(254, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(254, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(254, 5, xls.AddFormat(fmt));
            xls.SetCellValue(254, 5, "Sustrato de arena");
            xls.SetCellValue(254, 6, 630);

            fmt = xls.GetCellVisibleFormatDef(254, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(254, 7, xls.AddFormat(fmt));
            xls.SetCellValue(254, 7, "INPUTS\n(required amount * cost per unit)");

            fmt = xls.GetCellVisibleFormatDef(255, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(255, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(255, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(255, 5, xls.AddFormat(fmt));
            xls.SetCellValue(255, 5, "Sulfocalcio");
            xls.SetCellValue(255, 6, 0);

            fmt = xls.GetCellVisibleFormatDef(255, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(255, 7, xls.AddFormat(fmt));
            xls.SetCellValue(255, 7, "INPUTS\n(required amount * cost per unit)");

            fmt = xls.GetCellVisibleFormatDef(256, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(256, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(256, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(256, 5, xls.AddFormat(fmt));
            xls.SetCellValue(256, 5, "Cal");
            xls.SetCellValue(256, 6, 70);

            fmt = xls.GetCellVisibleFormatDef(256, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(256, 7, xls.AddFormat(fmt));
            xls.SetCellValue(256, 7, "INPUTS\n(required amount * cost per unit)");

            fmt = xls.GetCellVisibleFormatDef(257, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(257, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(257, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(257, 5, xls.AddFormat(fmt));
            xls.SetCellValue(257, 5, "Plastico");
            xls.SetCellValue(257, 6, 80);

            fmt = xls.GetCellVisibleFormatDef(257, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(257, 7, xls.AddFormat(fmt));
            xls.SetCellValue(257, 7, "INPUTS\n(required amount * cost per unit)");

            fmt = xls.GetCellVisibleFormatDef(258, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(258, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(258, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(258, 5, xls.AddFormat(fmt));
            xls.SetCellValue(258, 5, "Otros");
            xls.SetCellValue(258, 6, 1510);

            fmt = xls.GetCellVisibleFormatDef(258, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(258, 7, xls.AddFormat(fmt));
            xls.SetCellValue(258, 7, "INPUTS\n(required amount * cost per unit)");

            fmt = xls.GetCellVisibleFormatDef(259, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(259, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(259, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Style = TFlxFontStyles.Bold | TFlxFontStyles.Italic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(259, 5, xls.AddFormat(fmt));
            xls.SetCellValue(259, 5, "Materiales para Vivero o Almácigo (Año 0)");

            fmt = xls.GetCellVisibleFormatDef(259, 7);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            xls.SetCellFormat(259, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(260, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(260, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(260, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(260, 5, xls.AddFormat(fmt));
            xls.SetCellValue(260, 5, "Abono orgánico (Ej: Bocachi, otros)");
            xls.SetCellValue(260, 6, 2228);

            fmt = xls.GetCellVisibleFormatDef(260, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(260, 7, xls.AddFormat(fmt));
            xls.SetCellValue(260, 7, "INPUTS\n(required amount * cost per unit)");

            fmt = xls.GetCellVisibleFormatDef(261, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(261, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(261, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(261, 5, xls.AddFormat(fmt));
            xls.SetCellValue(261, 5, "Bolsitas de plastico");
            xls.SetCellValue(261, 6, 979.7);

            fmt = xls.GetCellVisibleFormatDef(261, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(261, 7, xls.AddFormat(fmt));
            xls.SetCellValue(261, 7, "INPUTS\n(required amount * cost per unit)");

            fmt = xls.GetCellVisibleFormatDef(262, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(262, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(262, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(262, 5, xls.AddFormat(fmt));
            xls.SetCellValue(262, 5, "Saran - Polisombra - Malla rache");
            xls.SetCellValue(262, 6, 1815);

            fmt = xls.GetCellVisibleFormatDef(262, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(262, 7, xls.AddFormat(fmt));
            xls.SetCellValue(262, 7, "INPUTS\n(required amount * cost per unit)");

            fmt = xls.GetCellVisibleFormatDef(263, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(263, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(263, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(263, 5, xls.AddFormat(fmt));
            xls.SetCellValue(263, 5, "Postes de madera");
            xls.SetCellValue(263, 6, 391);

            fmt = xls.GetCellVisibleFormatDef(263, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(263, 7, xls.AddFormat(fmt));
            xls.SetCellValue(263, 7, "INPUTS\n(required amount * cost per unit)");

            fmt = xls.GetCellVisibleFormatDef(264, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(264, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(264, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(264, 5, xls.AddFormat(fmt));
            xls.SetCellValue(264, 5, "Alambre de amarre");
            xls.SetCellValue(264, 6, 240);

            fmt = xls.GetCellVisibleFormatDef(264, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(264, 7, xls.AddFormat(fmt));
            xls.SetCellValue(264, 7, "INPUTS\n(required amount * cost per unit)");

            fmt = xls.GetCellVisibleFormatDef(265, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(265, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(265, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(265, 5, xls.AddFormat(fmt));
            xls.SetCellValue(265, 5, "Malla Ciclonica");
            xls.SetCellValue(265, 6, 1066);

            fmt = xls.GetCellVisibleFormatDef(265, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(265, 7, xls.AddFormat(fmt));
            xls.SetCellValue(265, 7, "INPUTS\n(required amount * cost per unit)");

            fmt = xls.GetCellVisibleFormatDef(266, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(266, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(266, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(266, 5, xls.AddFormat(fmt));
            xls.SetCellValue(266, 5, "Grapas");
            xls.SetCellValue(266, 6, 38.25);

            fmt = xls.GetCellVisibleFormatDef(266, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(266, 7, xls.AddFormat(fmt));
            xls.SetCellValue(266, 7, "INPUTS\n(required amount * cost per unit)");

            fmt = xls.GetCellVisibleFormatDef(267, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(267, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(267, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(267, 5, xls.AddFormat(fmt));
            xls.SetCellValue(267, 5, "Tierra para almacigos");
            xls.SetCellValue(267, 6, 3436);

            fmt = xls.GetCellVisibleFormatDef(267, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(267, 7, xls.AddFormat(fmt));
            xls.SetCellValue(267, 7, "INPUTS\n(required amount * cost per unit)");

            fmt = xls.GetCellVisibleFormatDef(268, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(268, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(268, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(268, 5, xls.AddFormat(fmt));
            xls.SetCellValue(268, 5, "Biofertilizantes líquidos (para foliar en el vivero)");
            xls.SetCellValue(268, 6, 482.5);

            fmt = xls.GetCellVisibleFormatDef(268, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(268, 7, xls.AddFormat(fmt));
            xls.SetCellValue(268, 7, "INPUTS\n(required amount * cost per unit)");

            fmt = xls.GetCellVisibleFormatDef(269, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(269, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(269, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(269, 5, xls.AddFormat(fmt));
            xls.SetCellValue(269, 5, "Agroquímicos (en el vivero)");
            xls.SetCellValue(269, 6, 0);

            fmt = xls.GetCellVisibleFormatDef(269, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(269, 7, xls.AddFormat(fmt));
            xls.SetCellValue(269, 7, "INPUTS\n(required amount * cost per unit)");

            fmt = xls.GetCellVisibleFormatDef(270, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(270, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(270, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(270, 5, xls.AddFormat(fmt));
            xls.SetCellValue(270, 5, "Fungicida");
            xls.SetCellValue(270, 6, 240);

            fmt = xls.GetCellVisibleFormatDef(270, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(270, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(271, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(271, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(271, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(271, 5, xls.AddFormat(fmt));
            xls.SetCellValue(271, 5, "Roca fosfórica");
            xls.SetCellValue(271, 6, 0);

            fmt = xls.GetCellVisibleFormatDef(271, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(271, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(272, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(272, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(272, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(272, 5, xls.AddFormat(fmt));
            xls.SetCellValue(272, 5, "Otros:");
            xls.SetCellValue(272, 6, 575.5);

            fmt = xls.GetCellVisibleFormatDef(272, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(272, 7, xls.AddFormat(fmt));
            xls.SetCellValue(272, 7, "INPUTS\n(required amount * cost per unit)");

            fmt = xls.GetCellVisibleFormatDef(273, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(273, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(273, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Style = TFlxFontStyles.Bold | TFlxFontStyles.Italic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(273, 5, xls.AddFormat(fmt));
            xls.SetCellValue(273, 5, "Materiales para Preparacion terreno y siembra (Año 0)");

            fmt = xls.GetCellVisibleFormatDef(273, 6);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(273, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(273, 7);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(273, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(274, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(274, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(274, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 320;
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Font.Family = 0;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0x00, 0x80, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            fmt.WrapText = true;
            xls.SetCellFormat(274, 5, xls.AddFormat(fmt));
            xls.SetCellValue(274, 5, "Abono orgánicos o COMPOST para LOS HOYOS");

            fmt = xls.GetCellVisibleFormatDef(274, 6);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xCC, 0xFF, 0xCC);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(274, 6, xls.AddFormat(fmt));
            xls.SetCellValue(274, 6, 3517.98883137063);

            fmt = xls.GetCellVisibleFormatDef(274, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(274, 7, xls.AddFormat(fmt));
            xls.SetCellValue(274, 7, "INPUTS\n(required amount * cost per unit)");

            fmt = xls.GetCellVisibleFormatDef(275, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(275, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(275, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.right;
            fmt.WrapText = true;
            xls.SetCellFormat(275, 5, xls.AddFormat(fmt));
            xls.SetCellValue(275, 5, "Harina de Roca");
            xls.SetCellValue(275, 6, 0);

            fmt = xls.GetCellVisibleFormatDef(275, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(275, 7, xls.AddFormat(fmt));
            xls.SetCellValue(275, 7, "INPUTS\n(required amount * cost per unit)");

            fmt = xls.GetCellVisibleFormatDef(276, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(276, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(276, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.right;
            fmt.WrapText = true;
            xls.SetCellFormat(276, 5, xls.AddFormat(fmt));
            xls.SetCellValue(276, 5, "Cascarilla de Café");
            xls.SetCellValue(276, 6, 0);

            fmt = xls.GetCellVisibleFormatDef(276, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(276, 7, xls.AddFormat(fmt));
            xls.SetCellValue(276, 7, "INPUTS\n(required amount * cost per unit)");

            fmt = xls.GetCellVisibleFormatDef(276, 8);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.right;
            fmt.WrapText = true;
            xls.SetCellFormat(276, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(277, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(277, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(277, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.right;
            fmt.WrapText = true;
            xls.SetCellFormat(277, 5, xls.AddFormat(fmt));
            xls.SetCellValue(277, 5, "Gallinaza");
            xls.SetCellValue(277, 6, 0);

            fmt = xls.GetCellVisibleFormatDef(277, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(277, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(277, 8);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.right;
            fmt.WrapText = true;
            xls.SetCellFormat(277, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(278, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(278, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(278, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 320;
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Font.Family = 0;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(278, 5, xls.AddFormat(fmt));
            xls.SetCellValue(278, 5, "Abono químico para los hoyos");

            fmt = xls.GetCellVisibleFormatDef(278, 6);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xCC, 0xFF, 0xCC);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(278, 6, xls.AddFormat(fmt));
            xls.SetCellValue(278, 6, 3517.98883137063);

            fmt = xls.GetCellVisibleFormatDef(278, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(278, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(278, 8);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.right;
            fmt.WrapText = true;
            xls.SetCellFormat(278, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(279, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(279, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(279, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(279, 5, xls.AddFormat(fmt));
            xls.SetCellValue(279, 5, "Cal");
            xls.SetCellValue(279, 6, 0);

            fmt = xls.GetCellVisibleFormatDef(279, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(279, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(279, 8);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.right;
            fmt.WrapText = true;
            xls.SetCellFormat(279, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(280, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(280, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(280, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(280, 5, xls.AddFormat(fmt));
            xls.SetCellValue(280, 5, "Otros elementos para los hoyos: ");
            xls.SetCellValue(280, 6, 0);

            fmt = xls.GetCellVisibleFormatDef(280, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(280, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(280, 8);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.right;
            fmt.WrapText = true;
            xls.SetCellFormat(280, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(281, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(281, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(281, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 300;
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Style = TFlxFontStyles.Bold | TFlxFontStyles.Italic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(281, 5, xls.AddFormat(fmt));

            Runs = new TRTFRun[2];
            Runs[0].FirstChar = 31;
            fnt = xls.GetDefaultFont;
            fnt.Name = "Arial";
            fnt.Size20 = 300;
            fnt.Color = TUIColor.FromArgb(0x00, 0x80, 0x00);
            fnt.Style = TFlxFontStyles.Bold | TFlxFontStyles.Italic;
            fnt.Underline = TFlxUnderline.Single;
            fnt.Scheme = TFontScheme.None;
            Runs[0].FontIndex = xls.AddFont(fnt);
            Runs[1].FirstChar = 39;
            fnt = xls.GetDefaultFont;
            fnt.Name = "Arial";
            fnt.Size20 = 300;
            fnt.Color = TExcelColor.Automatic;
            fnt.Style = TFlxFontStyles.Bold | TFlxFontStyles.Italic;
            fnt.Scheme = TFontScheme.None;
            Runs[1].FontIndex = xls.AddFont(fnt);
            xls.SetCellValue(281, 5, new TRichString("Materiales para levante en una hectárea (Año 1)", Runs, xls));
            //We could also have used: xls.SetCellFromHtml(281, 5, "Materiales para levante en una&nbsp;<font color = 'green'><u>hect&aacute;rea</u></font>&nbsp;(A&ntilde;o"
            //+" 1)")


    fmt = xls.GetCellVisibleFormatDef(281, 6);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(281, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(281, 7);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(281, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(281, 8);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.right;
            fmt.WrapText = true;
            xls.SetCellFormat(281, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(282, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(282, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(282, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 320;
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Font.Family = 0;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0x00, 0x80, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.WrapText = true;
            xls.SetCellFormat(282, 5, xls.AddFormat(fmt));
            xls.SetCellValue(282, 5, "Abono orgánicos o COMPOST  para levante");

            fmt = xls.GetCellVisibleFormatDef(282, 6);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xCC, 0xFF, 0xCC);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(282, 6, xls.AddFormat(fmt));
            xls.SetCellValue(282, 6, 1037.79389709306);

            fmt = xls.GetCellVisibleFormatDef(282, 7);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            xls.SetCellFormat(282, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(282, 8);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.right;
            fmt.WrapText = true;
            xls.SetCellFormat(282, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(283, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(283, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(283, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.right;
            fmt.WrapText = true;
            xls.SetCellFormat(283, 5, xls.AddFormat(fmt));
            xls.SetCellValue(283, 5, "Harina de Roca");
            xls.SetCellValue(283, 6, 0);

            fmt = xls.GetCellVisibleFormatDef(283, 7);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            xls.SetCellFormat(283, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(283, 8);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.right;
            fmt.WrapText = true;
            xls.SetCellFormat(283, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(284, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(284, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(284, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.right;
            fmt.WrapText = true;
            xls.SetCellFormat(284, 5, xls.AddFormat(fmt));
            xls.SetCellValue(284, 5, "Cascarilla de Café");
            xls.SetCellValue(284, 6, 0);

            fmt = xls.GetCellVisibleFormatDef(284, 7);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            xls.SetCellFormat(284, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(284, 8);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.right;
            fmt.WrapText = true;
            xls.SetCellFormat(284, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(285, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(285, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(285, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.right;
            fmt.WrapText = true;
            xls.SetCellFormat(285, 5, xls.AddFormat(fmt));
            xls.SetCellValue(285, 5, "Gallinaza");
            xls.SetCellValue(285, 6, 0);

            fmt = xls.GetCellVisibleFormatDef(285, 7);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            xls.SetCellFormat(285, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(285, 8);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.right;
            fmt.WrapText = true;
            xls.SetCellFormat(285, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(286, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(286, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(286, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 320;
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Font.Family = 0;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(286, 5, xls.AddFormat(fmt));
            xls.SetCellValue(286, 5, "Abono químico para levante (alrededor de la planta)");

            fmt = xls.GetCellVisibleFormatDef(286, 6);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xCC, 0xFF, 0xCC);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(286, 6, xls.AddFormat(fmt));
            xls.SetCellValue(286, 6, 1037.79389709306);

            fmt = xls.GetCellVisibleFormatDef(286, 7);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            xls.SetCellFormat(286, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(286, 8);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.right;
            fmt.WrapText = true;
            xls.SetCellFormat(286, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(287, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(287, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(287, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(287, 5, xls.AddFormat(fmt));
            xls.SetCellValue(287, 5, "Insumos para la foliación en la plantilla");
            xls.SetCellValue(287, 6, 0);

            fmt = xls.GetCellVisibleFormatDef(287, 7);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            xls.SetCellFormat(287, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(287, 8);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.right;
            fmt.WrapText = true;
            xls.SetCellFormat(287, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(288, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(288, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(288, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(288, 5, xls.AddFormat(fmt));
            xls.SetCellValue(288, 5, "Otros elementos para siembra y levante:");
            xls.SetCellValue(288, 6, 0);

            fmt = xls.GetCellVisibleFormatDef(288, 7);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            xls.SetCellFormat(288, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(288, 8);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.right;
            fmt.WrapText = true;
            xls.SetCellFormat(288, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(289, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(289, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(289, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 300;
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Style = TFlxFontStyles.Bold | TFlxFontStyles.Italic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.WrapText = true;
            xls.SetCellFormat(289, 5, xls.AddFormat(fmt));

            Runs = new TRTFRun[2];
            Runs[0].FirstChar = 29;
            fnt = xls.GetDefaultFont;
            fnt.Name = "Arial";
            fnt.Size20 = 300;
            fnt.Color = TUIColor.FromArgb(0x00, 0x80, 0x00);
            fnt.Style = TFlxFontStyles.Bold | TFlxFontStyles.Italic;
            fnt.Underline = TFlxUnderline.Single;
            fnt.Scheme = TFontScheme.None;
            Runs[0].FontIndex = xls.AddFont(fnt);
            Runs[1].FirstChar = 38;
            fnt = xls.GetDefaultFont;
            fnt.Name = "Arial";
            fnt.Size20 = 300;
            fnt.Color = TExcelColor.Automatic;
            fnt.Style = TFlxFontStyles.Bold | TFlxFontStyles.Italic;
            fnt.Scheme = TFontScheme.None;
            Runs[1].FontIndex = xls.AddFont(fnt);
            xls.SetCellValue(289, 5, new TRichString("Materiales para mantener una hectárea durante cosecha durante los años 2 a 8", Runs, xls));
            //We could also have used: xls.SetCellFromHtml(289, 5, "Materiales para mantener una&nbsp;<font color = 'green'><u>hect&aacute;rea&nbsp;</u></font>durante"
            //+" cosecha durante los a&ntilde;os 2 a 8")


    fmt = xls.GetCellVisibleFormatDef(289, 6);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(289, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(289, 7);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(289, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(290, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(290, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(290, 5);
            fmt.Font.Size20 = 320;
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0x00, 0x80, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.WrapText = true;
            xls.SetCellFormat(290, 5, xls.AddFormat(fmt));
            xls.SetCellValue(290, 5, "Abonos organicos mantenimiento");

            fmt = xls.GetCellVisibleFormatDef(290, 6);
            fmt.Font.Size20 = 400;
            fmt.Font.Family = 0;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent2, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(290, 6, xls.AddFormat(fmt));
            xls.SetCellValue(290, 6, new TFormula("='Inputs 1.0_metric_currency'!$D$25"));

            fmt = xls.GetCellVisibleFormatDef(290, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(290, 7, xls.AddFormat(fmt));
            xls.SetCellValue(290, 7, "INPUTS\n(required amount * cost per unit)");

            fmt = xls.GetCellVisibleFormatDef(290, 8);
            fmt.Font.Size20 = 400;
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            xls.SetCellFormat(290, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(291, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(291, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(291, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.right;
            fmt.WrapText = true;
            xls.SetCellFormat(291, 5, xls.AddFormat(fmt));
            xls.SetCellValue(291, 5, "Harina de Roca");
            xls.SetCellValue(291, 6, 0);

            fmt = xls.GetCellVisibleFormatDef(291, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(291, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(292, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(292, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(292, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.right;
            fmt.WrapText = true;
            xls.SetCellFormat(292, 5, xls.AddFormat(fmt));
            xls.SetCellValue(292, 5, "Cascarilla de Café");
            xls.SetCellValue(292, 6, 0);

            fmt = xls.GetCellVisibleFormatDef(292, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(292, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(293, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(293, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(293, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.right;
            fmt.WrapText = true;
            xls.SetCellFormat(293, 5, xls.AddFormat(fmt));
            xls.SetCellValue(293, 5, "Gallinaza");
            xls.SetCellValue(293, 6, 0);

            fmt = xls.GetCellVisibleFormatDef(293, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(293, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(294, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(294, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(294, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(294, 5, xls.AddFormat(fmt));
            xls.SetCellValue(294, 5, "Roca fosfórica");
            xls.SetCellValue(294, 6, 0);

            fmt = xls.GetCellVisibleFormatDef(294, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(294, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(295, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(295, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(295, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 320;
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Font.Family = 0;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.WrapText = true;
            xls.SetCellFormat(295, 5, xls.AddFormat(fmt));
            xls.SetCellValue(295, 5, "Abono químico para mantenimiento del cultivo ");

            fmt = xls.GetCellVisibleFormatDef(295, 6);
            fmt.Font.Size20 = 400;
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Family = 0;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent2, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(295, 6, xls.AddFormat(fmt));
            xls.SetCellValue(295, 6, new TFormula("='Inputs 1.0_metric_currency'!$D$24"));

            fmt = xls.GetCellVisibleFormatDef(295, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(295, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(296, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(296, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(296, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(296, 5, xls.AddFormat(fmt));
            xls.SetCellValue(296, 5, "Otro(s) abono (s):");
            xls.SetCellValue(296, 6, 0);

            fmt = xls.GetCellVisibleFormatDef(296, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(296, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(297, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(297, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(297, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 320;
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Font.Family = 0;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0x00, 0x80, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(297, 5, xls.AddFormat(fmt));
            xls.SetCellValue(297, 5, "Fertilizante organico para foliación:");

            fmt = xls.GetCellVisibleFormatDef(297, 6);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xCC, 0xFF, 0xCC);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(297, 6, xls.AddFormat(fmt));
            xls.SetCellValue(297, 6, 0);

            fmt = xls.GetCellVisibleFormatDef(297, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(297, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(298, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(298, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(298, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(298, 5, xls.AddFormat(fmt));
            xls.SetCellValue(298, 5, "Caldos bordeles");
            xls.SetCellValue(298, 6, 0);

            fmt = xls.GetCellVisibleFormatDef(298, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(298, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(299, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(299, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(299, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(299, 5, xls.AddFormat(fmt));
            xls.SetCellValue(299, 5, "Sulfocalcio");
            xls.SetCellValue(299, 6, 0);

            fmt = xls.GetCellVisibleFormatDef(299, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(299, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(300, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(300, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(300, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(300, 5, xls.AddFormat(fmt));
            xls.SetCellValue(300, 5, "Biofertilizante - multiminerales");
            xls.SetCellValue(300, 6, 0);

            fmt = xls.GetCellVisibleFormatDef(300, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(300, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(301, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(301, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(301, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 320;
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Font.Family = 0;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.left;
            xls.SetCellFormat(301, 5, xls.AddFormat(fmt));
            xls.SetCellValue(301, 5, "Químicos para foliación");

            fmt = xls.GetCellVisibleFormatDef(301, 6);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xCC, 0xFF, 0xCC);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(301, 6, xls.AddFormat(fmt));
            xls.SetCellValue(301, 6, 0);

            fmt = xls.GetCellVisibleFormatDef(301, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(301, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(302, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(302, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(302, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.left;
            xls.SetCellFormat(302, 5, xls.AddFormat(fmt));
            xls.SetCellValue(302, 5, "Otro(s) fertilizantes (s):");
            xls.SetCellValue(302, 6, 0);

            fmt = xls.GetCellVisibleFormatDef(302, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(302, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(303, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(303, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(303, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(303, 5, xls.AddFormat(fmt));
            xls.SetCellValue(303, 5, "Combustible:");
            xls.SetCellValue(303, 6, 0);

            fmt = xls.GetCellVisibleFormatDef(303, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(303, 7, xls.AddFormat(fmt));
            xls.SetCellValue(303, 7, "INPUTS\n(required amount * cost per unit)");

            fmt = xls.GetCellVisibleFormatDef(304, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(304, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(304, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(304, 5, xls.AddFormat(fmt));
            xls.SetCellValue(304, 5, "Otro(s) insumos para mantenimiento:");
            xls.SetCellValue(304, 6, 0);

            fmt = xls.GetCellVisibleFormatDef(304, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(304, 7, xls.AddFormat(fmt));
            xls.SetCellValue(304, 7, "INPUTS\n(required amount * cost per unit)");

            fmt = xls.GetCellVisibleFormatDef(305, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(305, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(305, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 320;
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(305, 5, xls.AddFormat(fmt));
            xls.SetCellValue(305, 5, "EQUIPOS Y MATERIALES REUTILIZABLES");

            fmt = xls.GetCellVisibleFormatDef(305, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 320;
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(305, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(305, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 320;
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            xls.SetCellFormat(305, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(305, 8);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 320;
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(305, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(305, 9);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 320;
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(305, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(306, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(306, 3, xls.AddFormat(fmt));
            xls.SetCellValue(306, 5, "Bomba manual ");
            xls.SetCellValue(306, 6, 1434);

            fmt = xls.GetCellVisibleFormatDef(306, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(306, 7, xls.AddFormat(fmt));
            xls.SetCellValue(306, 7, "INPUTS\n(required amount * cost per unit)");

            fmt = xls.GetCellVisibleFormatDef(307, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(307, 3, xls.AddFormat(fmt));
            xls.SetCellValue(307, 5, "Años de vida útil");

            fmt = xls.GetCellVisibleFormatDef(307, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(307, 6, xls.AddFormat(fmt));
            xls.SetCellValue(307, 6, 5.36);

            fmt = xls.GetCellVisibleFormatDef(307, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(307, 7, xls.AddFormat(fmt));
            xls.SetCellValue(307, 7, "INPUTS\n#");

            fmt = xls.GetCellVisibleFormatDef(308, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(308, 3, xls.AddFormat(fmt));
            xls.SetCellValue(308, 5, "Machete");
            xls.SetCellValue(308, 6, 340);

            fmt = xls.GetCellVisibleFormatDef(308, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(308, 7, xls.AddFormat(fmt));
            xls.SetCellValue(308, 7, "INPUTS\n(required amount * cost per unit)");

            fmt = xls.GetCellVisibleFormatDef(309, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(309, 3, xls.AddFormat(fmt));
            xls.SetCellValue(309, 5, "Años de vida útil");
            xls.SetCellValue(309, 6, 1.29);

            fmt = xls.GetCellVisibleFormatDef(309, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(309, 7, xls.AddFormat(fmt));
            xls.SetCellValue(309, 7, "INPUTS\n#");

            fmt = xls.GetCellVisibleFormatDef(310, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(310, 3, xls.AddFormat(fmt));
            xls.SetCellValue(310, 5, "Pala");
            xls.SetCellValue(310, 6, 184);

            fmt = xls.GetCellVisibleFormatDef(310, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(310, 7, xls.AddFormat(fmt));
            xls.SetCellValue(310, 7, "INPUTS\n(required amount * cost per unit)");

            fmt = xls.GetCellVisibleFormatDef(311, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(311, 3, xls.AddFormat(fmt));
            xls.SetCellValue(311, 5, "Años de vida útil");
            xls.SetCellValue(311, 6, 4.09);

            fmt = xls.GetCellVisibleFormatDef(311, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(311, 7, xls.AddFormat(fmt));
            xls.SetCellValue(311, 7, "INPUTS\n#");

            fmt = xls.GetCellVisibleFormatDef(312, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(312, 3, xls.AddFormat(fmt));
            xls.SetCellValue(312, 5, "Azadón");
            xls.SetCellValue(312, 6, 190);

            fmt = xls.GetCellVisibleFormatDef(312, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(312, 7, xls.AddFormat(fmt));
            xls.SetCellValue(312, 7, "INPUTS\n(required amount * cost per unit)");

            fmt = xls.GetCellVisibleFormatDef(313, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(313, 3, xls.AddFormat(fmt));
            xls.SetCellValue(313, 5, "Años de vida útil");
            xls.SetCellValue(313, 6, 4.8);

            fmt = xls.GetCellVisibleFormatDef(313, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(313, 7, xls.AddFormat(fmt));
            xls.SetCellValue(313, 7, "INPUTS\n#");

            fmt = xls.GetCellVisibleFormatDef(314, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(314, 3, xls.AddFormat(fmt));
            xls.SetCellValue(314, 5, "Carretilla");
            xls.SetCellValue(314, 6, 943);

            fmt = xls.GetCellVisibleFormatDef(314, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(314, 7, xls.AddFormat(fmt));
            xls.SetCellValue(314, 7, "INPUTS\n(required amount * cost per unit)");

            fmt = xls.GetCellVisibleFormatDef(315, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(315, 3, xls.AddFormat(fmt));
            xls.SetCellValue(315, 5, "Años de vida útil");
            xls.SetCellValue(315, 6, 4.84);

            fmt = xls.GetCellVisibleFormatDef(315, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(315, 7, xls.AddFormat(fmt));
            xls.SetCellValue(315, 7, "INPUTS\n#");

            fmt = xls.GetCellVisibleFormatDef(316, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(316, 3, xls.AddFormat(fmt));
            xls.SetCellValue(316, 5, "Lima");
            xls.SetCellValue(316, 6, 289);

            fmt = xls.GetCellVisibleFormatDef(316, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(316, 7, xls.AddFormat(fmt));
            xls.SetCellValue(316, 7, "INPUTS\n(required amount * cost per unit)");

            fmt = xls.GetCellVisibleFormatDef(317, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(317, 3, xls.AddFormat(fmt));
            xls.SetCellValue(317, 5, "Años de vida útil");
            xls.SetCellValue(317, 6, 0.63);

            fmt = xls.GetCellVisibleFormatDef(317, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(317, 7, xls.AddFormat(fmt));
            xls.SetCellValue(317, 7, "INPUTS\n#");

            fmt = xls.GetCellVisibleFormatDef(318, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(318, 3, xls.AddFormat(fmt));
            xls.SetCellValue(318, 5, "Chancha o ahoyador");
            xls.SetCellValue(318, 6, 210);

            fmt = xls.GetCellVisibleFormatDef(318, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(318, 7, xls.AddFormat(fmt));
            xls.SetCellValue(318, 7, "INPUTS\n(required amount * cost per unit)");

            fmt = xls.GetCellVisibleFormatDef(319, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(319, 3, xls.AddFormat(fmt));
            xls.SetCellValue(319, 5, "Años de vida útil");
            xls.SetCellValue(319, 6, 4.15);

            fmt = xls.GetCellVisibleFormatDef(319, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(319, 7, xls.AddFormat(fmt));
            xls.SetCellValue(319, 7, "INPUTS\n#");

            fmt = xls.GetCellVisibleFormatDef(320, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(320, 3, xls.AddFormat(fmt));
            xls.SetCellValue(320, 5, "Barretón");
            xls.SetCellValue(320, 6, 282);

            fmt = xls.GetCellVisibleFormatDef(320, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(320, 7, xls.AddFormat(fmt));
            xls.SetCellValue(320, 7, "INPUTS\n(required amount * cost per unit)");

            fmt = xls.GetCellVisibleFormatDef(321, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(321, 3, xls.AddFormat(fmt));
            xls.SetCellValue(321, 5, "Años de vida útil");
            xls.SetCellValue(321, 6, 5.04);

            fmt = xls.GetCellVisibleFormatDef(321, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(321, 7, xls.AddFormat(fmt));
            xls.SetCellValue(321, 7, "INPUTS\n#");

            fmt = xls.GetCellVisibleFormatDef(322, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(322, 3, xls.AddFormat(fmt));
            xls.SetCellValue(322, 5, "Mangueras");
            xls.SetCellValue(322, 6, 4409);

            fmt = xls.GetCellVisibleFormatDef(322, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(322, 7, xls.AddFormat(fmt));
            xls.SetCellValue(322, 7, "INPUTS\n(required amount * cost per unit)");

            fmt = xls.GetCellVisibleFormatDef(323, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(323, 3, xls.AddFormat(fmt));
            xls.SetCellValue(323, 5, "Años de vida útil");
            xls.SetCellValue(323, 6, 3.94);

            fmt = xls.GetCellVisibleFormatDef(323, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(323, 7, xls.AddFormat(fmt));
            xls.SetCellValue(323, 7, "INPUTS\n#");

            fmt = xls.GetCellVisibleFormatDef(324, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(324, 3, xls.AddFormat(fmt));
            xls.SetCellValue(324, 5, "Sistema de riego");
            xls.SetCellValue(324, 6, 203);

            fmt = xls.GetCellVisibleFormatDef(324, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(324, 7, xls.AddFormat(fmt));
            xls.SetCellValue(324, 7, "INPUTS\n(required amount * cost per unit)");

            fmt = xls.GetCellVisibleFormatDef(325, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(325, 3, xls.AddFormat(fmt));
            xls.SetCellValue(325, 5, "Años de vida útil");
            xls.SetCellValue(325, 6, 1.6);

            fmt = xls.GetCellVisibleFormatDef(325, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(325, 7, xls.AddFormat(fmt));
            xls.SetCellValue(325, 7, "INPUTS\n#");

            fmt = xls.GetCellVisibleFormatDef(326, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(326, 3, xls.AddFormat(fmt));
            xls.SetCellValue(326, 5, "Motosierra");
            xls.SetCellValue(326, 6, 8248);

            fmt = xls.GetCellVisibleFormatDef(326, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(326, 7, xls.AddFormat(fmt));
            xls.SetCellValue(326, 7, "INPUTS\n(required amount * cost per unit)");

            fmt = xls.GetCellVisibleFormatDef(327, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(327, 3, xls.AddFormat(fmt));
            xls.SetCellValue(327, 5, "Años de vida útil");
            xls.SetCellValue(327, 6, 7.03);

            fmt = xls.GetCellVisibleFormatDef(327, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(327, 7, xls.AddFormat(fmt));
            xls.SetCellValue(327, 7, "INPUTS\n#");

            fmt = xls.GetCellVisibleFormatDef(328, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(328, 3, xls.AddFormat(fmt));
            xls.SetCellValue(328, 5, "Serrucho");
            xls.SetCellValue(328, 6, 190);

            fmt = xls.GetCellVisibleFormatDef(328, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(328, 7, xls.AddFormat(fmt));
            xls.SetCellValue(328, 7, "INPUTS\n(required amount * cost per unit)");

            fmt = xls.GetCellVisibleFormatDef(329, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(329, 3, xls.AddFormat(fmt));
            xls.SetCellValue(329, 5, "Años de vida útil");
            xls.SetCellValue(329, 6, 3.77);

            fmt = xls.GetCellVisibleFormatDef(329, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(329, 7, xls.AddFormat(fmt));
            xls.SetCellValue(329, 7, "INPUTS\n#");

            fmt = xls.GetCellVisibleFormatDef(330, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(330, 3, xls.AddFormat(fmt));
            xls.SetCellValue(330, 5, "Bomba motor");
            xls.SetCellValue(330, 6, 0);

            fmt = xls.GetCellVisibleFormatDef(330, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(330, 7, xls.AddFormat(fmt));
            xls.SetCellValue(330, 7, "INPUTS\n(required amount * cost per unit)");

            fmt = xls.GetCellVisibleFormatDef(331, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(331, 3, xls.AddFormat(fmt));
            xls.SetCellValue(331, 5, "Años de vida útil");
            xls.SetCellValue(331, 6, 0);

            fmt = xls.GetCellVisibleFormatDef(331, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(331, 7, xls.AddFormat(fmt));
            xls.SetCellValue(331, 7, "INPUTS\n#");

            fmt = xls.GetCellVisibleFormatDef(332, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(332, 3, xls.AddFormat(fmt));
            xls.SetCellValue(332, 5, "Tijeras Podar");
            xls.SetCellValue(332, 6, 267);

            fmt = xls.GetCellVisibleFormatDef(332, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(332, 7, xls.AddFormat(fmt));
            xls.SetCellValue(332, 7, "INPUTS\n(required amount * cost per unit)");

            fmt = xls.GetCellVisibleFormatDef(333, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(333, 3, xls.AddFormat(fmt));
            xls.SetCellValue(333, 5, "Años de vida útil");
            xls.SetCellValue(333, 6, 4.6);

            fmt = xls.GetCellVisibleFormatDef(333, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(333, 7, xls.AddFormat(fmt));
            xls.SetCellValue(333, 7, "INPUTS\n#");

            fmt = xls.GetCellVisibleFormatDef(334, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(334, 3, xls.AddFormat(fmt));
            xls.SetCellValue(334, 5, "Hacha");
            xls.SetCellValue(334, 6, 251);

            fmt = xls.GetCellVisibleFormatDef(334, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(334, 7, xls.AddFormat(fmt));
            xls.SetCellValue(334, 7, "INPUTS\n(required amount * cost per unit)");

            fmt = xls.GetCellVisibleFormatDef(335, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(335, 3, xls.AddFormat(fmt));
            xls.SetCellValue(335, 5, "Años de vida útil");
            xls.SetCellValue(335, 6, 7.65);

            fmt = xls.GetCellVisibleFormatDef(335, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(335, 7, xls.AddFormat(fmt));
            xls.SetCellValue(335, 7, "INPUTS\n#");

            fmt = xls.GetCellVisibleFormatDef(336, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(336, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(336, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Style = TFlxFontStyles.Bold | TFlxFontStyles.Italic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(336, 5, xls.AddFormat(fmt));
            xls.SetCellValue(336, 5, "Equipos y Materiales para la cosecha y otras actividades");

            fmt = xls.GetCellVisibleFormatDef(336, 7);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            xls.SetCellFormat(336, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(337, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(337, 3, xls.AddFormat(fmt));
            xls.SetCellValue(337, 5, "Bascula o balanza");
            xls.SetCellValue(337, 6, 0);

            fmt = xls.GetCellVisibleFormatDef(337, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(337, 7, xls.AddFormat(fmt));
            xls.SetCellValue(337, 7, "INPUTS\n(required amount * cost per unit)");

            fmt = xls.GetCellVisibleFormatDef(338, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(338, 3, xls.AddFormat(fmt));
            xls.SetCellValue(338, 5, "Años de vida útil");
            xls.SetCellValue(338, 6, 8.14);

            fmt = xls.GetCellVisibleFormatDef(338, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(338, 7, xls.AddFormat(fmt));
            xls.SetCellValue(338, 7, "INPUTS\n#");

            fmt = xls.GetCellVisibleFormatDef(339, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(339, 3, xls.AddFormat(fmt));
            xls.SetCellValue(339, 5, "Vehiculo o automovil para trabajo");
            xls.SetCellValue(339, 6, 78478);

            fmt = xls.GetCellVisibleFormatDef(339, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(339, 7, xls.AddFormat(fmt));
            xls.SetCellValue(339, 7, "INPUTS\n(required amount * cost per unit)");

            fmt = xls.GetCellVisibleFormatDef(340, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(340, 3, xls.AddFormat(fmt));
            xls.SetCellValue(340, 5, "Años de vida útil");
            xls.SetCellValue(340, 6, 19);

            fmt = xls.GetCellVisibleFormatDef(340, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(340, 7, xls.AddFormat(fmt));
            xls.SetCellValue(340, 7, "INPUTS\n#");

            fmt = xls.GetCellVisibleFormatDef(341, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(341, 3, xls.AddFormat(fmt));
            xls.SetCellValue(341, 5, "Animal de trabajo");
            xls.SetCellValue(341, 6, 13471);

            fmt = xls.GetCellVisibleFormatDef(341, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(341, 7, xls.AddFormat(fmt));
            xls.SetCellValue(341, 7, "INPUTS\n(required amount * cost per unit)");

            fmt = xls.GetCellVisibleFormatDef(342, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(342, 3, xls.AddFormat(fmt));
            xls.SetCellValue(342, 5, "Años de vida útil");
            xls.SetCellValue(342, 6, 8.8);

            fmt = xls.GetCellVisibleFormatDef(342, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(342, 7, xls.AddFormat(fmt));
            xls.SetCellValue(342, 7, "INPUTS\n#");

            fmt = xls.GetCellVisibleFormatDef(343, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(343, 3, xls.AddFormat(fmt));
            xls.SetCellValue(343, 5, "Motocicleta");
            xls.SetCellValue(343, 6, 0);

            fmt = xls.GetCellVisibleFormatDef(343, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(343, 7, xls.AddFormat(fmt));
            xls.SetCellValue(343, 7, "INPUTS\n(required amount * cost per unit)");

            fmt = xls.GetCellVisibleFormatDef(344, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(344, 3, xls.AddFormat(fmt));
            xls.SetCellValue(344, 5, "Años de vida útil");
            xls.SetCellValue(344, 6, 0);

            fmt = xls.GetCellVisibleFormatDef(344, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(344, 7, xls.AddFormat(fmt));
            xls.SetCellValue(344, 7, "INPUTS\n#");

            fmt = xls.GetCellVisibleFormatDef(345, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(345, 3, xls.AddFormat(fmt));
            xls.SetCellValue(345, 5, "Sacos para la recoleccion");
            xls.SetCellValue(345, 6, 362);

            fmt = xls.GetCellVisibleFormatDef(345, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(345, 7, xls.AddFormat(fmt));
            xls.SetCellValue(345, 7, "INPUTS\n(required amount * cost per unit)");

            fmt = xls.GetCellVisibleFormatDef(346, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(346, 3, xls.AddFormat(fmt));
            xls.SetCellValue(346, 5, "Años de vida útil");
            xls.SetCellValue(346, 6, 1.3);

            fmt = xls.GetCellVisibleFormatDef(346, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(346, 7, xls.AddFormat(fmt));
            xls.SetCellValue(346, 7, "INPUTS\n#");

            fmt = xls.GetCellVisibleFormatDef(347, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(347, 3, xls.AddFormat(fmt));
            xls.SetCellValue(347, 5, "Sacos Pergamino");
            xls.SetCellValue(347, 6, 1328);

            fmt = xls.GetCellVisibleFormatDef(347, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(347, 7, xls.AddFormat(fmt));
            xls.SetCellValue(347, 7, "INPUTS\n(required amount * cost per unit)");

            fmt = xls.GetCellVisibleFormatDef(348, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(348, 3, xls.AddFormat(fmt));
            xls.SetCellValue(348, 5, "Años de vida útil");
            xls.SetCellValue(348, 6, 2.7);

            fmt = xls.GetCellVisibleFormatDef(348, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(348, 7, xls.AddFormat(fmt));
            xls.SetCellValue(348, 7, "INPUTS\n#");

            fmt = xls.GetCellVisibleFormatDef(349, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(349, 3, xls.AddFormat(fmt));
            xls.SetCellValue(349, 5, "Cabuya:");
            xls.SetCellValue(349, 6, 72.16);

            fmt = xls.GetCellVisibleFormatDef(349, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(349, 7, xls.AddFormat(fmt));
            xls.SetCellValue(349, 7, "INPUTS\n(required amount * cost per unit)");

            fmt = xls.GetCellVisibleFormatDef(350, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(350, 3, xls.AddFormat(fmt));
            xls.SetCellValue(350, 5, "Años de vida útil");
            xls.SetCellValue(350, 6, 1.03);

            fmt = xls.GetCellVisibleFormatDef(350, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(350, 7, xls.AddFormat(fmt));
            xls.SetCellValue(350, 7, "INPUTS\n#");

            fmt = xls.GetCellVisibleFormatDef(351, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(351, 3, xls.AddFormat(fmt));
            xls.SetCellValue(351, 5, "Canastas");
            xls.SetCellValue(351, 6, 302);

            fmt = xls.GetCellVisibleFormatDef(351, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(351, 7, xls.AddFormat(fmt));
            xls.SetCellValue(351, 7, "INPUTS\n(required amount * cost per unit)");

            fmt = xls.GetCellVisibleFormatDef(352, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(352, 3, xls.AddFormat(fmt));
            xls.SetCellValue(352, 5, "Años de vida útil");
            xls.SetCellValue(352, 6, 1.35);

            fmt = xls.GetCellVisibleFormatDef(352, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(352, 7, xls.AddFormat(fmt));
            xls.SetCellValue(352, 7, "INPUTS\n#");

            fmt = xls.GetCellVisibleFormatDef(353, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(353, 3, xls.AddFormat(fmt));
            xls.SetCellValue(353, 5, "Cajas");
            xls.SetCellValue(353, 6, 228);

            fmt = xls.GetCellVisibleFormatDef(353, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(353, 7, xls.AddFormat(fmt));
            xls.SetCellValue(353, 7, "INPUTS\n(required amount * cost per unit)");

            fmt = xls.GetCellVisibleFormatDef(354, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(354, 3, xls.AddFormat(fmt));
            xls.SetCellValue(354, 5, "Años de vida útil");
            xls.SetCellValue(354, 6, 6.3);

            fmt = xls.GetCellVisibleFormatDef(354, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(354, 7, xls.AddFormat(fmt));
            xls.SetCellValue(354, 7, "INPUTS\n#");

            fmt = xls.GetCellVisibleFormatDef(355, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(355, 3, xls.AddFormat(fmt));
            xls.SetCellValue(355, 5, "Otros");
            xls.SetCellValue(355, 6, 0);

            fmt = xls.GetCellVisibleFormatDef(355, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(355, 7, xls.AddFormat(fmt));
            xls.SetCellValue(355, 7, "INPUTS\n(required amount * cost per unit)");

            fmt = xls.GetCellVisibleFormatDef(356, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(356, 3, xls.AddFormat(fmt));
            xls.SetCellValue(356, 5, "Años de vida útil");
            xls.SetCellValue(356, 6, 1.9);

            fmt = xls.GetCellVisibleFormatDef(356, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(356, 7, xls.AddFormat(fmt));
            xls.SetCellValue(356, 7, "INPUTS\n#");

            fmt = xls.GetCellVisibleFormatDef(357, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(357, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(357, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Style = TFlxFontStyles.Bold | TFlxFontStyles.Italic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(357, 5, xls.AddFormat(fmt));
            xls.SetCellValue(357, 5, "Equipos y Materiales para el beneficio");

            fmt = xls.GetCellVisibleFormatDef(357, 7);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            xls.SetCellFormat(357, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(358, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(358, 3, xls.AddFormat(fmt));
            xls.SetCellValue(358, 5, "Despulpadora");
            xls.SetCellValue(358, 6, 7946);

            fmt = xls.GetCellVisibleFormatDef(358, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(358, 7, xls.AddFormat(fmt));
            xls.SetCellValue(358, 7, "INPUTS\n(required amount * cost per unit)");

            fmt = xls.GetCellVisibleFormatDef(359, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(359, 3, xls.AddFormat(fmt));
            xls.SetCellValue(359, 5, "Años de vida útil");
            xls.SetCellValue(359, 6, 7.5);

            fmt = xls.GetCellVisibleFormatDef(359, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(359, 7, xls.AddFormat(fmt));
            xls.SetCellValue(359, 7, "INPUTS\n#");

            fmt = xls.GetCellVisibleFormatDef(360, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(360, 3, xls.AddFormat(fmt));
            xls.SetCellValue(360, 5, "Sifon-Tolba");
            xls.SetCellValue(360, 6, 0);

            fmt = xls.GetCellVisibleFormatDef(360, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(360, 7, xls.AddFormat(fmt));
            xls.SetCellValue(360, 7, "INPUTS\n(required amount * cost per unit)");

            fmt = xls.GetCellVisibleFormatDef(361, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(361, 3, xls.AddFormat(fmt));
            xls.SetCellValue(361, 5, "Años de vida útil");
            xls.SetCellValue(361, 6, 0.1);

            fmt = xls.GetCellVisibleFormatDef(361, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(361, 7, xls.AddFormat(fmt));
            xls.SetCellValue(361, 7, "INPUTS\n#");

            fmt = xls.GetCellVisibleFormatDef(362, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(362, 3, xls.AddFormat(fmt));
            xls.SetCellValue(362, 5, "Motor");
            xls.SetCellValue(362, 6, 6565);

            fmt = xls.GetCellVisibleFormatDef(362, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(362, 7, xls.AddFormat(fmt));
            xls.SetCellValue(362, 7, "INPUTS\n(required amount * cost per unit)");

            fmt = xls.GetCellVisibleFormatDef(363, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(363, 3, xls.AddFormat(fmt));
            xls.SetCellValue(363, 5, "Años de vida útil");
            xls.SetCellValue(363, 6, 8.78);

            fmt = xls.GetCellVisibleFormatDef(363, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(363, 7, xls.AddFormat(fmt));
            xls.SetCellValue(363, 7, "INPUTS\n#");

            fmt = xls.GetCellVisibleFormatDef(364, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(364, 3, xls.AddFormat(fmt));
            xls.SetCellValue(364, 5, "Tanques o pilas de fermentacion");
            xls.SetCellValue(364, 6, 10236);

            fmt = xls.GetCellVisibleFormatDef(364, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(364, 7, xls.AddFormat(fmt));
            xls.SetCellValue(364, 7, "INPUTS\n(required amount * cost per unit)");

            fmt = xls.GetCellVisibleFormatDef(365, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(365, 3, xls.AddFormat(fmt));
            xls.SetCellValue(365, 5, "Años de vida útil");
            xls.SetCellValue(365, 6, 8.77);

            fmt = xls.GetCellVisibleFormatDef(365, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(365, 7, xls.AddFormat(fmt));
            xls.SetCellValue(365, 7, "INPUTS\n#");

            fmt = xls.GetCellVisibleFormatDef(366, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(366, 3, xls.AddFormat(fmt));
            xls.SetCellValue(366, 5, "Canal de correo para lavar café");
            xls.SetCellValue(366, 6, 2389);

            fmt = xls.GetCellVisibleFormatDef(366, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(366, 7, xls.AddFormat(fmt));
            xls.SetCellValue(366, 7, "INPUTS\n(required amount * cost per unit)");

            fmt = xls.GetCellVisibleFormatDef(367, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(367, 3, xls.AddFormat(fmt));
            xls.SetCellValue(367, 5, "Años de vida útil");
            xls.SetCellValue(367, 6, 7.47);

            fmt = xls.GetCellVisibleFormatDef(367, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(367, 7, xls.AddFormat(fmt));
            xls.SetCellValue(367, 7, "INPUTS\n#");

            fmt = xls.GetCellVisibleFormatDef(368, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(368, 3, xls.AddFormat(fmt));
            xls.SetCellValue(368, 5, "Tubos PVC");
            xls.SetCellValue(368, 6, 392);

            fmt = xls.GetCellVisibleFormatDef(368, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(368, 7, xls.AddFormat(fmt));
            xls.SetCellValue(368, 7, "INPUTS\n(required amount * cost per unit)");

            fmt = xls.GetCellVisibleFormatDef(369, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(369, 3, xls.AddFormat(fmt));
            xls.SetCellValue(369, 5, "Años de vida útil");
            xls.SetCellValue(369, 6, 5.16);

            fmt = xls.GetCellVisibleFormatDef(369, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(369, 7, xls.AddFormat(fmt));
            xls.SetCellValue(369, 7, "INPUTS\n#");

            fmt = xls.GetCellVisibleFormatDef(370, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(370, 3, xls.AddFormat(fmt));
            xls.SetCellValue(370, 5, "Sistema de filtración de agua (finca orgánica)");
            xls.SetCellValue(370, 6, 530);

            fmt = xls.GetCellVisibleFormatDef(370, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(370, 7, xls.AddFormat(fmt));
            xls.SetCellValue(370, 7, "INPUTS\n(required amount * cost per unit)");

            fmt = xls.GetCellVisibleFormatDef(371, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(371, 3, xls.AddFormat(fmt));
            xls.SetCellValue(371, 5, "Años de vida útil");
            xls.SetCellValue(371, 6, 6.13);

            fmt = xls.GetCellVisibleFormatDef(371, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(371, 7, xls.AddFormat(fmt));
            xls.SetCellValue(371, 7, "INPUTS\n#");

            fmt = xls.GetCellVisibleFormatDef(372, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(372, 3, xls.AddFormat(fmt));
            xls.SetCellValue(372, 5, "Criba - Zaranda");
            xls.SetCellValue(372, 6, 227);

            fmt = xls.GetCellVisibleFormatDef(372, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(372, 7, xls.AddFormat(fmt));
            xls.SetCellValue(372, 7, "INPUTS\n(required amount * cost per unit)");

            fmt = xls.GetCellVisibleFormatDef(373, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(373, 3, xls.AddFormat(fmt));
            xls.SetCellValue(373, 5, "Años de vida útil");
            xls.SetCellValue(373, 6, 5.3);

            fmt = xls.GetCellVisibleFormatDef(373, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(373, 7, xls.AddFormat(fmt));
            xls.SetCellValue(373, 7, "INPUTS\n#");

            fmt = xls.GetCellVisibleFormatDef(374, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(374, 3, xls.AddFormat(fmt));
            xls.SetCellValue(374, 5, "Desmucilagador");
            xls.SetCellValue(374, 6, 0);

            fmt = xls.GetCellVisibleFormatDef(374, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(374, 7, xls.AddFormat(fmt));
            xls.SetCellValue(374, 7, "INPUTS\n(required amount * cost per unit)");

            fmt = xls.GetCellVisibleFormatDef(375, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(375, 3, xls.AddFormat(fmt));
            xls.SetCellValue(375, 5, "Años de vida útil");
            xls.SetCellValue(375, 6, 0);

            fmt = xls.GetCellVisibleFormatDef(375, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(375, 7, xls.AddFormat(fmt));
            xls.SetCellValue(375, 7, "INPUTS\n#");

            fmt = xls.GetCellVisibleFormatDef(376, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(376, 3, xls.AddFormat(fmt));
            xls.SetCellValue(376, 5, "Pozo");
            xls.SetCellValue(376, 6, 442);

            fmt = xls.GetCellVisibleFormatDef(376, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(376, 7, xls.AddFormat(fmt));
            xls.SetCellValue(376, 7, "INPUTS\n(required amount * cost per unit)");

            fmt = xls.GetCellVisibleFormatDef(377, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(377, 3, xls.AddFormat(fmt));
            xls.SetCellValue(377, 5, "Años de vida útil");
            xls.SetCellValue(377, 6, 9.5);

            fmt = xls.GetCellVisibleFormatDef(377, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(377, 7, xls.AddFormat(fmt));
            xls.SetCellValue(377, 7, "INPUTS\n#");

            fmt = xls.GetCellVisibleFormatDef(378, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(378, 3, xls.AddFormat(fmt));
            xls.SetCellValue(378, 5, "Otro componente del beneficio húmedo");
            xls.SetCellValue(378, 6, 0);

            fmt = xls.GetCellVisibleFormatDef(378, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(378, 7, xls.AddFormat(fmt));
            xls.SetCellValue(378, 7, "INPUTS\n(required amount * cost per unit)");

            fmt = xls.GetCellVisibleFormatDef(379, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(379, 3, xls.AddFormat(fmt));
            xls.SetCellValue(379, 5, "Años de vida útil");
            xls.SetCellValue(379, 6, 0.1);

            fmt = xls.GetCellVisibleFormatDef(379, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(379, 7, xls.AddFormat(fmt));
            xls.SetCellValue(379, 7, "INPUTS\n#");

            fmt = xls.GetCellVisibleFormatDef(380, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(380, 3, xls.AddFormat(fmt));
            xls.SetCellValue(380, 5, "Secador solar - Plancha concreto");
            xls.SetCellValue(380, 6, 25522);

            fmt = xls.GetCellVisibleFormatDef(380, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(380, 7, xls.AddFormat(fmt));
            xls.SetCellValue(380, 7, "INPUTS\n(required amount * cost per unit)");

            fmt = xls.GetCellVisibleFormatDef(381, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(381, 3, xls.AddFormat(fmt));
            xls.SetCellValue(381, 5, "Años de vida útil");
            xls.SetCellValue(381, 6, 8.3);

            fmt = xls.GetCellVisibleFormatDef(381, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(381, 7, xls.AddFormat(fmt));
            xls.SetCellValue(381, 7, "INPUTS\n#");

            fmt = xls.GetCellVisibleFormatDef(382, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(382, 3, xls.AddFormat(fmt));
            xls.SetCellValue(382, 5, "Plastico");
            xls.SetCellValue(382, 6, 1521);

            fmt = xls.GetCellVisibleFormatDef(382, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(382, 7, xls.AddFormat(fmt));
            xls.SetCellValue(382, 7, "INPUTS\n(required amount * cost per unit)");

            fmt = xls.GetCellVisibleFormatDef(383, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(383, 3, xls.AddFormat(fmt));
            xls.SetCellValue(383, 5, "Años de vida útil");
            xls.SetCellValue(383, 6, 3.43);

            fmt = xls.GetCellVisibleFormatDef(383, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(383, 7, xls.AddFormat(fmt));
            xls.SetCellValue(383, 7, "INPUTS\n#");

            fmt = xls.GetCellVisibleFormatDef(384, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(384, 3, xls.AddFormat(fmt));
            xls.SetCellValue(384, 5, "Rastrillo");
            xls.SetCellValue(384, 6, 228);

            fmt = xls.GetCellVisibleFormatDef(384, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(384, 7, xls.AddFormat(fmt));
            xls.SetCellValue(384, 7, "INPUTS\n(required amount * cost per unit)");

            fmt = xls.GetCellVisibleFormatDef(385, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(385, 3, xls.AddFormat(fmt));
            xls.SetCellValue(385, 5, "Años de vida útil");
            xls.SetCellValue(385, 6, 2.91);

            fmt = xls.GetCellVisibleFormatDef(385, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(385, 7, xls.AddFormat(fmt));
            xls.SetCellValue(385, 7, "INPUTS\n#");

            fmt = xls.GetCellVisibleFormatDef(386, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(386, 3, xls.AddFormat(fmt));
            xls.SetCellValue(386, 5, "Escoba");
            xls.SetCellValue(386, 6, 50);

            fmt = xls.GetCellVisibleFormatDef(386, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(386, 7, xls.AddFormat(fmt));
            xls.SetCellValue(386, 7, "INPUTS\n(required amount * cost per unit)");

            fmt = xls.GetCellVisibleFormatDef(387, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(387, 3, xls.AddFormat(fmt));
            xls.SetCellValue(387, 5, "Años de vida útil");
            xls.SetCellValue(387, 6, 1.4);

            fmt = xls.GetCellVisibleFormatDef(387, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(387, 7, xls.AddFormat(fmt));
            xls.SetCellValue(387, 7, "INPUTS\n#");

            fmt = xls.GetCellVisibleFormatDef(388, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(388, 3, xls.AddFormat(fmt));
            xls.SetCellValue(388, 5, "Bodega");
            xls.SetCellValue(388, 6, 0);

            fmt = xls.GetCellVisibleFormatDef(388, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(388, 7, xls.AddFormat(fmt));
            xls.SetCellValue(388, 7, "INPUTS\n(required amount * cost per unit)");

            fmt = xls.GetCellVisibleFormatDef(389, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(389, 3, xls.AddFormat(fmt));
            xls.SetCellValue(389, 5, "Años de vida útil");
            xls.SetCellValue(389, 6, 0.1);

            fmt = xls.GetCellVisibleFormatDef(389, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(389, 7, xls.AddFormat(fmt));
            xls.SetCellValue(389, 7, "INPUTS\n#");

            fmt = xls.GetCellVisibleFormatDef(390, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(390, 3, xls.AddFormat(fmt));
            xls.SetCellValue(390, 5, "Otro componente del beneficio seco");
            xls.SetCellValue(390, 6, 75);

            fmt = xls.GetCellVisibleFormatDef(390, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(390, 7, xls.AddFormat(fmt));
            xls.SetCellValue(390, 7, "INPUTS\n(required amount * cost per unit)");

            fmt = xls.GetCellVisibleFormatDef(391, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(391, 3, xls.AddFormat(fmt));
            xls.SetCellValue(391, 5, "Años de vida útil");
            xls.SetCellValue(391, 6, 1.5);

            fmt = xls.GetCellVisibleFormatDef(391, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(391, 7, xls.AddFormat(fmt));
            xls.SetCellValue(391, 7, "INPUTS\n#");

            fmt = xls.GetCellVisibleFormatDef(392, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(392, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(392, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Style = TFlxFontStyles.Bold | TFlxFontStyles.Italic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(392, 5, xls.AddFormat(fmt));
            xls.SetCellValue(392, 5, "Otros materiales para el beneficio");

            fmt = xls.GetCellVisibleFormatDef(392, 7);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            xls.SetCellFormat(392, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(393, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(393, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(393, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.WrapText = true;
            xls.SetCellFormat(393, 5, xls.AddFormat(fmt));
            xls.SetCellValue(393, 5, "Que tantos litros de agua se pueden gastar en el beneficio húmedo de un KILO de café"
            + " pergamino seco?");
            xls.SetCellValue(393, 6, 348);

            fmt = xls.GetCellVisibleFormatDef(393, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(393, 7, xls.AddFormat(fmt));
            xls.SetCellValue(393, 7, "INPUTS\n#");

            fmt = xls.GetCellVisibleFormatDef(393, 8);
            fmt.WrapText = true;
            xls.SetCellFormat(393, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(394, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(394, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(394, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.WrapText = true;
            xls.SetCellFormat(394, 5, xls.AddFormat(fmt));
            xls.SetCellValue(394, 5, "Cuánto paga por un litro de agua utilizado en su proceso de producción de café?");
            xls.SetCellValue(394, 6, 1);

            fmt = xls.GetCellVisibleFormatDef(394, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(394, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(395, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(395, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(395, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.WrapText = true;
            xls.SetCellFormat(395, 5, xls.AddFormat(fmt));
            xls.SetCellValue(395, 5, "Pago de energía utilizada únicamente pare el proceso de producción de café ");
            xls.SetCellValue(395, 6, 0);

            fmt = xls.GetCellVisibleFormatDef(395, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(395, 7, xls.AddFormat(fmt));
            xls.SetCellValue(395, 7, "INPUTS\n#");

            fmt = xls.GetCellVisibleFormatDef(396, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(396, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(396, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.WrapText = true;
            fmt.Indent = 1;
            xls.SetCellFormat(396, 5, xls.AddFormat(fmt));
            xls.SetCellValue(396, 5, "Otros materiales para el beneficio:");
            xls.SetCellValue(396, 6, 0);

            fmt = xls.GetCellVisibleFormatDef(396, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(396, 7, xls.AddFormat(fmt));
            xls.SetCellValue(396, 7, "INPUTS\n#");

            fmt = xls.GetCellVisibleFormatDef(397, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(397, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(397, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 320;
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.left;
            xls.SetCellFormat(397, 5, xls.AddFormat(fmt));
            xls.SetCellValue(397, 5, "Costos relacionados a la asociación");

            fmt = xls.GetCellVisibleFormatDef(397, 7);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            xls.SetCellFormat(397, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(398, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(398, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(398, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(398, 5, xls.AddFormat(fmt));
            xls.SetCellValue(398, 5, "Costo de entrada / Costo de inscripción");
            xls.SetCellValue(398, 6, 2236);

            fmt = xls.GetCellVisibleFormatDef(398, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(398, 7, xls.AddFormat(fmt));
            xls.SetCellValue(398, 7, "INPUTS\n#");

            fmt = xls.GetCellVisibleFormatDef(399, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(399, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(399, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(399, 5, xls.AddFormat(fmt));
            xls.SetCellValue(399, 5, "Ahorro / Fondo rotatorio (anual)");
            xls.SetCellValue(399, 6, 0);

            fmt = xls.GetCellVisibleFormatDef(399, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(399, 7, xls.AddFormat(fmt));
            xls.SetCellValue(399, 7, "INPUTS\n#");

            fmt = xls.GetCellVisibleFormatDef(400, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(400, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(400, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(400, 5, xls.AddFormat(fmt));
            xls.SetCellValue(400, 5, "Seguro de vida anual");
            xls.SetCellValue(400, 6, 0);

            fmt = xls.GetCellVisibleFormatDef(400, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(400, 7, xls.AddFormat(fmt));
            xls.SetCellValue(400, 7, "INPUTS\n#");

            fmt = xls.GetCellVisibleFormatDef(401, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(401, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(401, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 320;
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(401, 5, xls.AddFormat(fmt));
            xls.SetCellValue(401, 5, "Valor de la Tierra");

            fmt = xls.GetCellVisibleFormatDef(401, 7);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            xls.SetCellFormat(401, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(402, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(402, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(402, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(402, 5, xls.AddFormat(fmt));
            xls.SetCellValue(402, 5, "Es usted dueño de su tierra o es propiedad comunal?");
            xls.SetCellValue(402, 6, 1);

            fmt = xls.GetCellVisibleFormatDef(402, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(402, 7, xls.AddFormat(fmt));
            xls.SetCellValue(402, 7, "INPUTS\n(yes 1, no = 0)");

            fmt = xls.GetCellVisibleFormatDef(403, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(403, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(403, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(403, 5, xls.AddFormat(fmt));

            Runs = new TRTFRun[1];
            Runs[0].FirstChar = 19;
            fnt = xls.GetDefaultFont;
            fnt.Name = "Arial";
            fnt.Size20 = 280;
            fnt.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fnt.Style = TFlxFontStyles.Bold;
            fnt.Scheme = TFontScheme.None;
            Runs[0].FontIndex = xls.AddFont(fnt);
            xls.SetCellValue(403, 5, new TRichString("Valor por hectarea (DE LA TIERRA POR SI SOLA, SIN CULTIVO)", Runs, xls));
            //We could also have used: xls.SetCellFromHtml(403, 5, "Valor por hectarea&nbsp;<font color = 'blue'><b>(DE LA TIERRA POR SI SOLA, SIN CULTIVO)</b></font>")

            xls.SetCellValue(403, 6, 36217);

            fmt = xls.GetCellVisibleFormatDef(403, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(403, 7, xls.AddFormat(fmt));
            xls.SetCellValue(403, 7, "INPUTS\n#");

            fmt = xls.GetCellVisibleFormatDef(404, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(404, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(404, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.WrapText = true;
            xls.SetCellFormat(404, 5, xls.AddFormat(fmt));

            Runs = new TRTFRun[2];
            Runs[0].FirstChar = 68;
            fnt = xls.GetDefaultFont;
            fnt.Name = "Arial";
            fnt.Size20 = 280;
            fnt.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fnt.Style = TFlxFontStyles.Bold;
            fnt.Scheme = TFontScheme.None;
            Runs[0].FontIndex = xls.AddFont(fnt);
            Runs[1].FirstChar = 73;
            fnt = xls.GetDefaultFont;
            fnt.Name = "Arial";
            fnt.Size20 = 280;
            fnt.Color = TExcelColor.Automatic;
            fnt.Scheme = TFontScheme.None;
            Runs[1].FontIndex = xls.AddFont(fnt);
            xls.SetCellValue(404, 5, new TRichString("En caso de no ser propietario,  paga alguna renta, cuál es el valor ANUAL?", Runs, xls));
            //We could also have used: xls.SetCellFromHtml(404, 5, "En caso de no ser propietario, &nbsp;paga alguna renta, cu&aacute;l es el valor&nbsp;<font"
            //+" color = 'blue'><b>ANUAL</b></font>?")

    xls.SetCellValue(404, 6, 0);

            fmt = xls.GetCellVisibleFormatDef(404, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(404, 7, xls.AddFormat(fmt));
            xls.SetCellValue(404, 7, "INPUTS\n#");

            fmt = xls.GetCellVisibleFormatDef(405, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(405, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(405, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 320;
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(405, 5, xls.AddFormat(fmt));
            xls.SetCellValue(405, 5, "Impuestos y Regulación");

            fmt = xls.GetCellVisibleFormatDef(405, 7);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            xls.SetCellFormat(405, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(406, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(406, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(406, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.left;
            xls.SetCellFormat(406, 5, xls.AddFormat(fmt));
            xls.SetCellValue(406, 5, "Impuesto a la propiedad en PESOS (Catastro)");
            xls.SetCellValue(406, 6, "147..28");

            fmt = xls.GetCellVisibleFormatDef(406, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(406, 7, xls.AddFormat(fmt));
            xls.SetCellValue(406, 7, "INPUTS\n#");

            fmt = xls.GetCellVisibleFormatDef(407, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(407, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(407, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.left;
            xls.SetCellFormat(407, 5, xls.AddFormat(fmt));
            xls.SetCellValue(407, 5, "Impuestos no oficiales en PESOS");
            xls.SetCellValue(407, 6, 0);

            fmt = xls.GetCellVisibleFormatDef(407, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(407, 7, xls.AddFormat(fmt));
            xls.SetCellValue(407, 7, "INPUTS\n#");

            fmt = xls.GetCellVisibleFormatDef(408, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(408, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(408, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(408, 5, xls.AddFormat(fmt));
            xls.SetCellValue(408, 5, "Otros:");
            xls.SetCellValue(408, 6, 0);

            fmt = xls.GetCellVisibleFormatDef(408, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(408, 7, xls.AddFormat(fmt));
            xls.SetCellValue(408, 7, "INPUTS\n#");

            fmt = xls.GetCellVisibleFormatDef(409, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(409, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(409, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 320;
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(409, 5, xls.AddFormat(fmt));
            xls.SetCellValue(409, 5, "Costos Administrativos e imprevistos");

            fmt = xls.GetCellVisibleFormatDef(409, 7);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            xls.SetCellFormat(409, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(410, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(410, 3, xls.AddFormat(fmt));
            xls.SetCellValue(410, 4, "sólo horas totales");

            fmt = xls.GetCellVisibleFormatDef(410, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.WrapText = true;
            xls.SetCellFormat(410, 5, xls.AddFormat(fmt));
            xls.SetCellValue(410, 5, "Cuánto tiempo puede gastar  supervisando (no trabajando) actividades como limpias,"
            + " manejos, podas, obras conservación, cosecha etc");
            xls.SetCellValue(410, 6, 14);

            fmt = xls.GetCellVisibleFormatDef(410, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(410, 7, xls.AddFormat(fmt));
            xls.SetCellValue(410, 7, "INPUTS\n#");

            fmt = xls.GetCellVisibleFormatDef(411, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(411, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(411, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.WrapText = true;
            xls.SetCellFormat(411, 5, xls.AddFormat(fmt));
            xls.SetCellValue(411, 5, "Cuánto tiempo puede gastar  en cuestiones administrativas de su finca tales como llevar"
            + " las cuentas, los registros, pagar servicios, pagar trabajdores, ir al banco, ir a"
            + " la asociación por papeles, pagos, reuniones (NO capacitaciones).");
            xls.SetCellValue(411, 6, 5.37);

            fmt = xls.GetCellVisibleFormatDef(411, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(411, 7, xls.AddFormat(fmt));
            xls.SetCellValue(411, 7, "INPUTS\n#");

            fmt = xls.GetCellVisibleFormatDef(412, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(412, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(412, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.WrapText = true;
            xls.SetCellFormat(412, 5, xls.AddFormat(fmt));
            xls.SetCellValue(412, 5, "Cuánto tiempo puede gastar  en capacitar a la gente que contrata para las diversas"
            + " labores de la finca");
            xls.SetCellValue(412, 6, 1.4);

            fmt = xls.GetCellVisibleFormatDef(412, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(412, 7, xls.AddFormat(fmt));
            xls.SetCellValue(412, 7, "INPUTS\n#");

            fmt = xls.GetCellVisibleFormatDef(413, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(413, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(413, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.WrapText = true;
            xls.SetCellFormat(413, 5, xls.AddFormat(fmt));
            xls.SetCellValue(413, 5, "Cuánto puede gastar  En costos extraordinarios tales como cubrir asistencias médicas"
            + " por accidentes de trabajo de sus trabajadores");
            xls.SetCellValue(413, 6, 899);

            fmt = xls.GetCellVisibleFormatDef(413, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(413, 7, xls.AddFormat(fmt));
            xls.SetCellValue(413, 7, "INPUTS\n#");

            fmt = xls.GetCellVisibleFormatDef(414, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(414, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(414, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 320;
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(414, 5, xls.AddFormat(fmt));
            xls.SetCellValue(414, 5, "Costos no mencionados");
            xls.SetCellValue(414, 6, 0);

            fmt = xls.GetCellVisibleFormatDef(414, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(414, 7, xls.AddFormat(fmt));
            xls.SetCellValue(414, 7, "INPUTS\n#");

            fmt = xls.GetCellVisibleFormatDef(415, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(415, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(415, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 320;
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(415, 5, xls.AddFormat(fmt));
            xls.SetCellValue(415, 5, "Costos no mencionados");
            xls.SetCellValue(415, 6, 0);

            fmt = xls.GetCellVisibleFormatDef(415, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(415, 7, xls.AddFormat(fmt));
            xls.SetCellValue(415, 7, "INPUTS\n#");

            fmt = xls.GetCellVisibleFormatDef(416, 3);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(416, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(416, 4);
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            xls.SetCellFormat(416, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(416, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 320;
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            xls.SetCellFormat(416, 5, xls.AddFormat(fmt));
            xls.SetCellValue(416, 5, "Costos no mencionados");

            fmt = xls.GetCellVisibleFormatDef(416, 6);
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            xls.SetCellFormat(416, 6, xls.AddFormat(fmt));
            xls.SetCellValue(416, 6, 24);

            fmt = xls.GetCellVisibleFormatDef(416, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(416, 7, xls.AddFormat(fmt));
            xls.SetCellValue(416, 7, "INPUTS\n#");

            //Comments

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
            xls.SetComment(301, 6, new TRichString("Juan Hernandez:\n976 para que cuadre con el link externo?", Runs, xls));

            //You probably don't need to call the lines below. This code is needed only if you want to change the comment box properties like color or default location
            TCommentProperties CommentProps = TCommentProperties.CreateStandard(301, 6, xls);
            CommentProps.Anchor = new TClientAnchor(TFlxAnchorType.DontMoveAndDontResize, 293, 64, 7, 47, 296, 26, 7, 805);

            //Excel by doesn't autofit the comment box so it can hold all text.
            //There is an option in TCommentProperties, but if you use it Excel will show the text in a single line.
            //To have FlexCel autofit the comment for you, you can do it with the following code:

            //    CommentProps.Anchor = xls.AutofitComment(new TRichString("Juan Hernandez:\n976 para que cuadre con el link externo?", Runs, xls), 1.5, true, 1.1, 0, CommentProps.Anchor);

            xls.SetCommentProperties(301, 6, CommentProps);

            //Objects
            TShapeProperties ShapeOptions1 = new TShapeProperties();
            ShapeOptions1.Anchor = new TClientAnchor(TFlxAnchorType.MoveAndResize, 3, 79, 8, 377, 4, 236, 16, 391);
            ShapeOptions1.ShapeType = TShapeType.Rectangle;
            ShapeOptions1.ObjectType = TObjectType.MicrosoftOfficeDrawing;
            ShapeOptions1.ShapeName = "TextBox 1";
            ShapeOptions1.Text = "Nota: Subieron costos de producción porque sumé \n\nCosto de aimentar un trabajador"
            + " por día + costo del jornal\n\nPorque no se estaba considerando antes el costo de"
            + " alimentación en esta edición";
            ShapeOptions1.TextFlags = 530;
            ShapeOptions1.RotateTextWithShape = true;
            ShapeOptions1.ShapeThemeFont = new TShapeFont(TFontScheme.Minor, TDrawingColor.FromTheme(TThemeColor.Foreground1));
            ShapeOptions1.Print = true;
            ShapeOptions1.Visible = true;
            ShapeOptions1.ShapeGeometry = "<?xml version=\"1.0\" encoding=\"utf-8\" standalone=\"yes\"?><a:shapeGeom xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"><a:prstGeom"
            + " prst=\"rect\"><a:avLst /></a:prstGeom></a:shapeGeom>";
            ShapeOptions1.ShapeOptions.SetValue(TShapeOption.fillColor, 16777215);
            ShapeOptions1.ShapeOptions.SetValue(TShapeOption.fillBackColor, 134217808);
            ShapeOptions1.ShapeOptions.SetValue(TShapeOption.fFilled, true);
            ShapeOptions1.ShapeOptions.SetValue(TShapeOption.lineColor, 12369084);
            ShapeOptions1.ShapeOptions.SetValue(TShapeOption.shadowColor, 0);
            ShapeOptions1.ShapeOptions.SetValue(TShapeOption.wzName, "TextBox 1");
            xls.AddAutoShape(ShapeOptions1);


            //Cell selection and scroll position.
            xls.SelectCell(294, 8, false);
            xls.ScrollWindow(268, 3);

            //Standard Document Properties - Most are only for xlsx files. In xls files FlexCel will only change the Creation Date and Modified Date.
            xls.DocumentProperties.SetStandardProperty(TPropertyId.Author, "Mary Kate");

            //You will normally not set LastSavedBy, since this is a new file.
            //If you don't set it, FlexCel will use the creator instead.
            //    xls.DocumentProperties.SetStandardProperty(TPropertyId.LastSavedBy, "Juan Hernandez");


            //You will normally not set CreateDateTime, since this is a new file and FlexCel will automatically use the current datetime.
            //But if you are editing a file and want to preserve the original creation date, you need to either set PreserveCreationDate to true:
            //    xls.DocumentProperties.PreserveCreationDate = true;
            //Or you can hardcode a creating date by setting it in UTC time, ISO8601 format:
            //    xls.DocumentProperties.SetStandardProperty(TPropertyId.CreateTimeDate, "2015-01-07T22:31:31Z");


        }

    }
}
