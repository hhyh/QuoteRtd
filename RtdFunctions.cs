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
            string idx = "0";

            // China security
            if (Regex.Match(name, "^s[hz][0-9]{6}$").Success)
                switch(item)
                {
                    case "name":
                        idx = "0";
                        break;
                    case "latest":
                        idx = "3";
                        break;
                    case "date":
                        idx = "30";
                        break;
                    case "time":
                        idx = "31";
                        break;
                    case "high":
                        idx = "4";
                        break;
                    case "low":
                        idx = "5";
                        break;
                    case "open":
                        idx = "1";
                        break;
                    case "close":
                        idx = "2";
                        break;
                    case "volume":
                        idx = "8";
                        break;
                    case "amount":
                        idx = "9";
                        break;
                }
            // US security
            else if (Regex.Match(name, "^gb_\\w+$").Success)
                switch (item)
                {
                    case "name":
                        idx = "0";
                        break;
                    case "latest":
                        idx = "1";
                        break;
                    case "date":
                        idx = "3";
                        break;
                    case "time":
                        idx = "25";
                        break;
                    case "high":
                        idx = "6";
                        break;
                    case "low":
                        idx = "7";
                        break;
                    case "open":
                        idx = "5";
                        break;
                    case "close":
                        idx = "26";
                        break;
                    case "volume":
                        idx = "10";
                        break;
                }
            // HongKong security
            else if (Regex.Match(name, "^hk[0-9]{5}$").Success ||
                     Regex.Match(name, "^hk\\w+$").Success)
                switch (item)
                {
                    case "name":
                        idx = "1";
                        break;
                    case "latest":
                        idx = "6";
                        break;
                    case "date":
                        idx = "17";
                        break;
                    case "time":
                        idx = "18";
                        break;
                    case "high":
                        idx = "4";
                        break;
                    case "low":
                        idx = "5";
                        break;
                    case "open":
                        idx = "2";
                        break;
                    case "close":
                        idx = "3";
                        break;
                    case "volume":
                        idx = "12";
                        break;
                    case "amount":
                        idx = "11";
                        break;
                }
            // public fund
            else if (Regex.Match(name, "^of[0-9]{6}$").Success)
                switch (item)
                {
                    case "name":
                        idx = "0";
                        break;
                    case "latest":
                        idx = "1";
                        break;
                    case "date":
                        idx = "5";
                        break;
                    case "close":
                        idx = "3";
                        break;
                }
            // forex
            else if (Regex.Match(name, "^fx_\\w+$").Success)
                switch (item)
                {
                    case "name":
                        idx = "9";
                        break;
                    case "latest":
                        idx = "1";
                        break;
                    case "date":
                        idx = "17";
                        break;
                    case "time":
                        idx = "0";
                        break;
                    case "high":
                        idx = "6";
                        break;
                    case "low":
                        idx = "7";
                        break;
                    case "open":
                        idx = "5";
                        break;
                    case "close":
                        idx = "3";
                        break;
                }
            // China stock future
            else if (Regex.Match(name, "^nf_I[CHF][0-9]{4}$").Success)
                switch (item)
                {
                    case "name":
                        idx = "49";
                        break;
                    case "latest":
                        idx = "3";
                        break;
                    case "date":
                        idx = "36";
                        break;
                    case "time":
                        idx = "37";
                        break;
                    case "high":
                        idx = "1";
                        break;
                    case "low":
                        idx = "2";
                        break;
                    case "open":
                        idx = "0";
                        break;
                    case "close":
                        idx = "13";
                        break;
                }

            return XlCall.RTD("QuoteRtd.QuoteServer", null, name, idx, item);
        }
    }
}
