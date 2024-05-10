using NSEUtilitaire;
using System;

namespace BOD_Utility
{
    public class FTPCRED
    {
        public string Username { get; set; }
        public string Password { get; set; }
    }

    public class ContractMasterData
    {
        public string Symbol { get; set; }
        public string InstName { get; set; }
        public string ScripType { get; set; }
        public string StrikePrice { get; set; }
        public DateTime ExpiryDate { get; set; }
    }

    public class OTMFileData
    {
        public string Symbol { get; set; }
        public string InstName { get; set; }
        public string ScripType { get; set; }
        public string StrikePrice { get; set; }
        public DateTime ExpiryDate { get; set; }
        public string Percentage { get; set; }
    }

    public class NiftyOTMFile
    {
        public string Symbol { get; set; }
        public string ExpiryDate { get; set; }
        public string OTMPercentage { get; set; } = "0";
        public string OTHPercentage { get; set; } = "0";
    }


    public class _CDBhavcopy
    {
        public string Symbol { get; internal set; }

        /// <summary>
        /// Symbol|Expiry|Strike|ScripType (NIFTY|29OCT2020|11500.00|CE OR NIFTY|29OCT2020|0|XX) 
        /// </summary>
        public string CustomScripname { get; internal set; }
        public en_Instrument Instrument { get; internal set; }

        public DateTime Expiry { get; internal set; }
        public double ExpiryUnix { get; internal set; } = 0;
        /// <summary>
        /// 0 for XX.
        /// </summary>
        public double StrikePrice { get; internal set; }
        public en_ScripType ScripType { get; internal set; }

        public double Open { get; internal set; }
        public double High { get; internal set; }
        public double Low { get; internal set; }
        public double Close { get; internal set; }
        public double PreviousClose { get; internal set; }

        public long QtyTraded { get; internal set; }

        public double ValueInLacs { get; internal set; }

        public long OpenInterest { get; internal set; }
        public long ChangeInOpenInterest { get; internal set; }
    }

}
