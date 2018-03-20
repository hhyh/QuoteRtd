using System;
using System.Collections.Generic;
using System.Text;
using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;

namespace QuoteRtd
{
    class ExcelMacros : IExcelAddIn
    {
        public void AutoClose()
        {
        }

        public void AutoOpen()
        {
            Application app = ExcelDnaUtil.Application as Application;
            app.SheetCalculate += new AppEvents_SheetCalculateEventHandler(SheetCalc);
        }

        private void SheetCalc(object sh)
        {
            Worksheet sheet = sh as Worksheet;
            foreach (Range r in sheet.UsedRange)
            {
                string text = r.Text as string;

                if (text.EndsWith("%"))
                {
                    if ((double)r.Value > 0)
                    {
                        r.Font.ColorIndex = 3;
                    }
                    else if ((double)r.Value < 0)
                    {
                        r.Font.ColorIndex = 10;
                    }
                    else
                    {
                        r.Font.Color = XlColorIndex.xlColorIndexAutomatic;
                    }
                }
            }
        }

    }
}
