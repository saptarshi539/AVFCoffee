using FlexCel.Core;
using FlexCel.XlsAdapter;
using System;
using System.Collections.Generic;
using System.Text;

namespace Domain.Flexcel
{
    public class Flexcelsum
    {
        public String sumcells()
        {
            XlsFile xls = new XlsFile(1, TExcelFileFormat.v2016, true);

            //Enters a string into A1.

            xls.SetCellValue(1, 1, "Hello from FlexCel!");

            //Enters a number into A2.
            //Note that xls.SetCellValue(2, 1, "7") would enter a string.
            xls.SetCellValue(2, 1, 7);

            //Enter another floating point number.
            //All numbers in Excel are floating point,
            //so even if you enter an integer, it will be stored as double.
            xls.SetCellValue(3, 1, 11.3);

            //Enters a formula into A4.
            xls.SetCellValue(4, 1, new TFormula("=Sum(A2:A3)"));
            
            //Saves the file to the "Documents" folder.
            xls.Save(System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Personal), "test.xlsx"));

            return Convert.ToString(xls.GetCellValue(4,1));
        }
    }
}
