using System;
using System.Collections.Generic;
using System.Text;
using ExcelDna.Integration;
using System.Runtime.InteropServices;
using ExcelDna.Integration.CustomUI;
using ExcelDna.Logging;

namespace QuoteRtd
{
    [ComVisible(true)]
    public class QuoteRibbon : ExcelRibbon
    {
        public bool Checkbox_getPressed(IRibbonControl control)
        {
            switch (control.Id)
            {
                case "AutoUpdate":
                    return GlobalConfig.refreshData;
                case "Debug":
                    return GlobalConfig.logEnable;
            }
            return false;
        }

        public void Checkbox_onAction(IRibbonControl control, bool pressed)
        {
            switch (control.Id)
            {
                case "AutoUpdate":
                    GlobalConfig.refreshData = pressed;
                    break;
                case "Debug":
                    GlobalConfig.logEnable = pressed;
                    break;
            }
        }

        public void Editbox_onChange(IRibbonControl control, string text)
        {
            GlobalConfig.refreshInterval = int.Parse(text);
            if (GlobalConfig.refreshInterval < 2000)
            {
                LogDisplay.WriteLine("Refresh interval too small, reset to 2000.");
                GlobalConfig.refreshInterval = 2000;
            }
            else
                LogDisplay.WriteLine("Refresh interval set to " + text + ".");
        }

        public string Editbox_getText(IRibbonControl control)
        {
            return GlobalConfig.refreshInterval.ToString();
        }
    }
}
