using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using ExcelDna.Integration;

namespace QuoteRtd
{
    // Help function to simplify the RTD call
    public static class RtdFunctions
    {
        [ExcelFunction(Name = "Quote")]
        public static object Quote(string name, string item)
        {
            // Map the item name to item index

            // China security
            if (Regex.Match(name, "^s[hz][0-9]{6}$").Success)
                switch(item)
                {
                    case "name":
                        item = "0";
                        break;
                    case "latest":
                        item = "3";
                        break;
                    case "date":
                        item = "30";
                        break;
                    case "time":
                        item = "31";
                        break;
                    case "high":
                        item = "4";
                        break;
                    case "low":
                        item = "5";
                        break;
                    case "open":
                        item = "1";
                        break;
                    case "close":
                        item = "2";
                        break;
                    case "volume":
                        item = "8";
                        break;
                    case "amount":
                        item = "9";
                        break;
                }
            // US security
            else if (Regex.Match(name, "^gb_\\w+$").Success)
                switch (item)
                {
                    case "name":
                        item = "0";
                        break;
                    case "latest":
                        item = "1";
                        break;
                    case "date":
                        item = "3";
                        break;
                    case "time":
                        item = "25";
                        break;
                    case "high":
                        item = "6";
                        break;
                    case "low":
                        item = "7";
                        break;
                    case "open":
                        item = "5";
                        break;
                    case "close":
                        item = "26";
                        break;
                    case "volume":
                        item = "10";
                        break;
                }
            // HongKong security
            else if (Regex.Match(name, "^hk[0-9]{5}$").Success)
                switch (item)
                {
                    case "name":
                        item = "1";
                        break;
                    case "latest":
                        item = "6";
                        break;
                    case "date":
                        item = "17";
                        break;
                    case "time":
                        item = "18";
                        break;
                    case "high":
                        item = "4";
                        break;
                    case "low":
                        item = "5";
                        break;
                    case "open":
                        item = "2";
                        break;
                    case "close":
                        item = "3";
                        break;
                    case "volume":
                        item = "12";
                        break;
                    case "amount":
                        item = "11";
                        break;
                }

            return XlCall.RTD("QuoteRtd.QuoteServer", null, name, item);
        }
    }
}
