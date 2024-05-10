using n.Structs;
using NerveLog;
using NSEUtilitaire;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace BOD_Utility
{
    internal static class Day1
    {
        //added on 09APR2021 by Amey. To notify user if there is inconsistency in Day1 file.
        internal static bool isAnyError = false;
        internal static string ApplicationPath = Application.StartupPath + "\\";
        internal static NerveLogger _logger = new NerveLogger(true, true, ApplicationName: "BOD-Utility");
        internal static string Filename = "";

        public static List<EODPositionInfo> Read(string DAY1Folder, string BhavcopyPath, NerveLogger _logger, bool MarkToClosing, bool UseClosing, HashSet<string> hs_Usernames,
            ConcurrentDictionary<string, ContractMaster> dict_ScripInfo, ConcurrentDictionary<string, ContractMaster> dict_CustomScripInfo,
            ConcurrentDictionary<string, ContractMaster> dict_TokenScripInfo)
        {
            isAnyError = false;

            List<EODPositionInfo> list_Day1Positions = new List<EODPositionInfo>();
            HashSet<string> hs_Index = new HashSet<string>();   //Added by Akshay on 26-03-2021

            try
            {
                EODPositionInfo _EODPositionInfo;
                ConcurrentDictionary<string, string> dict_ClientFamily = new ConcurrentDictionary<string, string>();

                DateTime dte_ScripExpiry;
                double ExpiryInTicks;

                var directory = new DirectoryInfo("C:/Prime/Day1");

                //added on 30OCT2020 by Amey
                var BhavcopyDirectory = new DirectoryInfo("C:/Prime");
                var Day1Directory = new DirectoryInfo("C:/Prime/Day1");

                var FOBhavcopy = BhavcopyDirectory.GetFiles("NSE_FO_bhavcopy*.csv")
                               .OrderByDescending(f => f.LastWriteTime)
                               .First();

                var CMBhavcopy = BhavcopyDirectory.GetFiles("NSE_CM_bhavcopy_*.csv")
                           .OrderByDescending(f => f.LastWriteTime)
                           .First();

                var list_FOBhavcopy = Exchange.ReadFOBhavcopy(FOBhavcopy.FullName,true);
                var list_CMBhavcopy = Exchange.ReadCMBhavcopy(CMBhavcopy.FullName, true);

                ContractMaster ScripInfo = new ContractMaster();
                string CustomScripNameKey = string.Empty;

                var Day1File = Day1Directory.GetFiles()
                                              .OrderByDescending(f => f.LastWriteTime)
                                              .First();

                var arr_Day1 = File.ReadAllLines(Day1File.FullName);


                string Symbol = string.Empty;
                string ScripType = string.Empty;
                double StrikePrice = -0;
                string CustomScripName = string.Empty;
                string ScripNameKey = string.Empty;
                string newScripNameKey = string.Empty;
                long NetQty = 0;
                double NetValue = 0;
                double Price = 0;
                double AvgPrice = 0;
                n.Structs.en_Segment Segment = n.Structs.en_Segment.NSEFO;


                foreach (var _Line in arr_Day1)
                {
                    var arr_Fields = _Line.Split(',').Select(v => v.Trim()).ToArray();
                    try
                    {
                        //added for testing
                        //hs_Usernames.Add(arr_Fields[62].ToUpper());

                        var _Username = arr_Fields[0].ToUpper();

                        //changed from AccountID to LoginID on 23MAR2021 by Amey
                        //if (!hs_Usernames.Contains(arr_Fields[3].ToUpper())) continue;

                        //changed from AccountID to LoginID on 23MAR2021 by Amey
                        if (!hs_Usernames.Contains(_Username))
                        {
                            continue;
                        }

                        NetQty = Convert.ToInt64(Convert.ToDouble(arr_Fields[8]));
                        if (NetQty == 0) { continue; }

                        Segment = (n.Structs.en_Segment)Enum.Parse(typeof(n.Structs.en_Segment), arr_Fields[1]);
                        ScripType = arr_Fields[5];

                        if (ScripType == "") continue;

                        if (Segment == n.Structs.en_Segment.NSECM || Segment == n.Structs.en_Segment.BSECM)
                            dte_ScripExpiry = Convert.ToDateTime("01JAN1980");
                        else
                            dte_ScripExpiry = Convert.ToDateTime(arr_Fields[3]);

                        ExpiryInTicks = ConvertToUnixTimestamp(dte_ScripExpiry);

                        //changed on 27MAR2021 by Amey
                        Symbol = arr_Fields[2].ToUpper();
                        if (Symbol == "") continue;

                        if (ExpiryInTicks == 0 || ScripType == "XX")
                            StrikePrice = 0;
                        else
                            StrikePrice = Convert.ToDouble(arr_Fields[4]);

                        AvgPrice = Convert.ToDouble(arr_Fields[9]);

                        //if (ExpiryInTicks == 0)
                        //    Segment = en_Segment.NSECM;
                        //else
                        //    Segment = en_Segment.NSEFO;

                        //NetValue = //Math.Abs(Convert.ToDouble(arr_Fields[13]) - Convert.ToDouble(arr_Fields[15]));

                        //if (NetQty == 0 || (dte_ScripExpiry.Date < DateTime.Now.Date && ExpiryInTicks != 0))
                        //{
                        //    AvgPrice = 0;
                        //}
                        //else
                        //    AvgPrice = Convert.ToDouble(NetValue / Math.Abs(NetQty));


                        CustomScripName = $"{Symbol}|{dte_ScripExpiry.ToString("ddMMMyyyy").ToUpper()}|{(StrikePrice == 0 ? "0" : StrikePrice.ToString("#.00"))}|{ScripType}";

                        CustomScripNameKey = Segment + "|" + CustomScripName;

                        if (!dict_CustomScripInfo.TryGetValue(CustomScripNameKey, out ScripInfo))
                        {
                            if (!dict_TokenScripInfo.TryGetValue(Segment + "|" + arr_Fields[6], out ScripInfo))
                            {
                                _logger.Debug("Read Day1 EntrySkipped | " + _Line);
                                continue;
                            }
                        }
                        //Old MarkToClosing code is commented for New Prime
                        //if (!dict_ScripInfo.TryGetValue(ScripNameKey, out ScripInfo))
                        //{
                        //    if (!dict_TokenScripInfo.TryGetValue(Segment + "|" + arr_Fields[6], out ScripInfo))
                        //    {
                        //        _logger.WriteLog("Read Day1 EntrySkipped | " + _Line);
                        //        continue;
                        //    }
                        //}
                        //#region OldCode
                        ////added on 29OCT2020 by Amey
                        //if (MarkToClosing)
                        //{
                        //    try
                        //    {
                        //        if (Segment == en_Segment.NSECM)
                        //        {
                        //            //added by omkar for Snehadri-New-Primes
                        //            if (dict_CustomScripInfo.TryGetValue(newScripNameKey, out ContractMaster _ScripInfo))
                        //            {
                        //                if (UseClosing)
                        //                    AvgPrice = _ScripInfo.ClosingPrice;
                        //                else
                        //                    AvgPrice = _ScripInfo.SettlementPrice;
                        //            }
                        //            else
                        //                AvgPrice = Convert.ToDouble(arr_Fields[9]);
                        //            //--

                        //            //var ClosePrice = list_CMBhavcopy.Where(v => v.CustomScripname.Equals(ScripInfo.CustomScripName)).FirstOrDefault();
                        //            //if (ClosePrice is null)
                        //            //    _logger.WriteLog($"Closing Not Found For : {CustomScripName}", true);
                        //            //else
                        //            //    AvgPrice = ClosePrice.Close;
                        //        }
                        //        else if (Segment == en_Segment.BSECM)
                        //        {
                        //            //added by omkar for Snehadri-New-Primes
                        //            if (dict_CustomScripInfo.TryGetValue(newScripNameKey, out ContractMaster _ScripInfo))
                        //            {
                        //                if (UseClosing)
                        //                    AvgPrice = _ScripInfo.ClosingPrice;
                        //                else
                        //                    AvgPrice = _ScripInfo.SettlementPrice;
                        //            }
                        //            else
                        //                AvgPrice = Convert.ToDouble(arr_Fields[9]);
                        //            //--

                        //            //if (dict_BseCMBhavcopy.TryGetValue(ScripInfo.Token, out double ClosePrice))
                        //            //{
                        //            //    AvgPrice = ClosePrice;
                        //            //}
                        //            //else
                        //            //    _logger.WriteLog($"Closing Not Found For : {CustomScripName}", true);
                        //        }
                        //        else
                        //        {
                        //            //added by omkar for Snehadri-New-Primes
                        //            if (dict_ScripInfo.TryGetValue(ScripNameKey, out ContractMaster _ScripInfo))
                        //            {
                        //                if (UseClosing)
                        //                    AvgPrice = _ScripInfo.ClosingPrice;
                        //                else
                        //                    AvgPrice = _ScripInfo.SettlementPrice;
                        //            }
                        //            else
                        //                AvgPrice = Convert.ToDouble(arr_Fields[9]);
                        //            //--

                        //                // var ClosePrice = list_FOBhavcopy.Where(v => v.CustomScripname.Equals(ScripInfo.CustomScripName)).FirstOrDefault();

                        //                //if (ClosePrice is null)
                        //                //    _logger.WriteLog($"Closing Not Found For : {CustomScripName}", true);
                        //                //else if (UseClosing)
                        //                //    AvgPrice = ClosePrice.Close;
                        //                //else
                        //                //    AvgPrice = ClosePrice.SettlePrice;
                        //        }
                        //    }
                        //    catch (Exception ee) { _logger.WriteLog("Day1 Loop Closing : " + _Line + Environment.NewLine + ee); }
                        //}
                        //#endregion


                        #region Snehadri-New-Primes
                        try
                        {
                            if (dict_CustomScripInfo.TryGetValue(CustomScripNameKey, out ContractMaster _ScripInfo))
                            {
                                if (MarkToClosing)
                                {
                                    if (UseClosing)
                                        Price = _ScripInfo.ClosingPrice;
                                    else
                                        Price = _ScripInfo.SettlementPrice;
                                }
                                else
                                    Price = Math.Abs(Math.Round((Convert.ToDouble(arr_Fields[13]) - Convert.ToDouble(arr_Fields[15])) / NetQty, 2));
                            }
                            else
                            {
                                _logger.Debug($"CustomScripNameKey Not Found For : {CustomScripNameKey}");
                                continue;
                            }
                        }
                        catch (Exception ee) { _logger.Error(ee, "Day1 Loop Closing : " + _Line + Environment.NewLine + ee); }
                        #endregion


                        ////added on 20JAN2021 by Amey
                        //if (dte_ScripExpiry.Date < DateTime.Now.Date && ExpiryInTicks != 0)
                        //    continue;

                        //added on 20APR2021 by Amey


                        if (hs_Usernames.Contains(_Username))
                        {
                            _EODPositionInfo = new EODPositionInfo()
                            {
                                Username = arr_Fields[0].ToUpper(),
                                Segment = Segment,
                                Token = ScripInfo.Token,
                                TradePrice = Price,
                                TradeQuantity = NetQty,
                                UnderlyingSegment = ScripInfo.UnderlyingSegment,
                                UnderlyingToken = ScripInfo.UnderlyingToken
                            };

                            list_Day1Positions.Add(_EODPositionInfo);
                        }


                    }
                    catch (Exception ee) { _logger.Error(ee, "Day1 Loop : " + _Line + Environment.NewLine + ee); isAnyError = true; }
                }

            }
            catch (Exception ee)
            {
                _logger.Error(ee, "Read DAYFOCM " + ee);

                isAnyError = true;
            }

            return list_Day1Positions;
        }


        public static List<EODPositionInfo> ReadPS03(string PS03Folder, string BhavcopyPath, NerveLogger _logger, bool MarkToClosing, bool UseClosing, HashSet<string> hs_Usernames,
            ConcurrentDictionary<string, ContractMaster> dict_ScripInfo, ConcurrentDictionary<string, ContractMaster> dict_CustomScripInfo,
            ConcurrentDictionary<string, ContractMaster> dict_TokenScripInfo)
        {
            isAnyError = false;

            List<EODPositionInfo> list_Day1Positions = new List<EODPositionInfo>();

            try
            {
                EODPositionInfo _EODPositionInfo;

                var directory = new DirectoryInfo(PS03Folder);

                var PSO3File = directory.GetFiles().OrderByDescending(f => f.LastWriteTime).First();
                //added on 30OCT2020 by Amey
                var BhavcopyDirectory = new DirectoryInfo("C:/Prime");

                var FOBhavcopy = BhavcopyDirectory.GetFiles("NSE_FO_bhavcopy*.csv")
                           .OrderByDescending(f => f.LastWriteTime)
                           .First();

                var CMBhavcopy = BhavcopyDirectory.GetFiles("NSE_CM_bhavcopy_*.csv")
                           .OrderByDescending(f => f.LastWriteTime)
                           .First();

                var list_FOBhavcopy = Exchange.ReadFOBhavcopy(FOBhavcopy.FullName,true);
                var list_CMBhavcopy = Exchange.ReadCMBhavcopy(CMBhavcopy.FullName, true);

                using (FileStream stream = File.Open(PSO3File.FullName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    using (StreamReader sr = new StreamReader(stream))
                    {
                        string line1;

                        string Underlying = string.Empty;
                        double StrikePrice = 0;
                        string ScripType = string.Empty;
                        string CustomScripName = string.Empty;
                        string UnderlyingScripName = string.Empty;

                        //added on 20APR2021 by Amey
                        var CustomScripNameKey = string.Empty;
                        var Segment = n.Structs.en_Segment.NSEFO;
                        long Qty = 0;
                        double TradePrice = 0;
                        ContractMaster ScripInfo = new ContractMaster();

                        while ((line1 = sr.ReadLine()) != null)
                        {
                            string[] fields = line1.Split(',');

                            if (fields[0].Trim() != "")
                            {
                                try
                                {
                                    if (!hs_Usernames.Contains(fields[7].Trim().ToUpper())) continue;//added by Navin on 02-12-2019 to pick records of uploaded clients

                                    //added on 9SEP2020 to avoid 0 Qty Uploads
                                    Qty = Convert.ToInt64(Convert.ToDouble(fields[28].Trim())) - Convert.ToInt64(Convert.ToDouble(fields[30].Trim()));
                                    if (Qty == 0)
                                    {
                                        _logger.Debug("Data Incorrect PS03 Skipped Invalid Qty : " + Qty + Environment.NewLine + line1);
                                        continue;
                                    }

                                    Underlying = fields[9].Trim().ToUpper();
                                    ScripType = fields[12].Trim().ToUpper();

                                    if (ScripType == "FF")
                                        ScripType = "FUT";

                                    DateTime dte_ScripExpiry = DateTime.Parse(fields[10].Trim().ToUpper());

                                    if (dte_ScripExpiry.Date < DateTime.Now.Date && ScripType != "EQ")
                                    {
                                        _logger.Debug("Data Incorrect PS03 Skipped Expired : " + dte_ScripExpiry + "|" + ScripType + Environment.NewLine + line1);
                                        continue;
                                    }

                                    try { StrikePrice = Convert.ToDouble(fields[11]); } catch (Exception) { }

                                    CustomScripName = $"{Underlying}|{dte_ScripExpiry.ToString("ddMMMyyyy").ToUpper()}|{(StrikePrice == 0 ? "0" : StrikePrice.ToString("#.00"))}|{(ScripType == "FUT" ? "XX" : ScripType)}";

                                    //added on 20APR2021 by Amey
                                    CustomScripNameKey = Segment + "|" + CustomScripName;
                                    if (!dict_CustomScripInfo.TryGetValue(CustomScripNameKey, out ScripInfo))
                                        continue;

                                    //Snehadri-New-Primes
                                    try
                                    {
                                        if (dict_CustomScripInfo.TryGetValue(CustomScripNameKey, out ContractMaster _ScripInfo))
                                        {
                                            if (UseClosing)
                                                TradePrice = _ScripInfo.ClosingPrice;
                                            else
                                                TradePrice = _ScripInfo.SettlementPrice;
                                        }
                                    }
                                    catch (Exception ee) { _logger.Error(ee, "PS03 Loop Closing : " + line1 + Environment.NewLine + ee); }

                                    #region old code
                                    //try
                                    //{

                                    //    if (ScripType == "EQ")
                                    //    {
                                    //        var ClosePrice = list_CMBhavcopy.Where(v => v.CustomScripname.Equals(CustomScripName)).FirstOrDefault();
                                    //        if (ClosePrice is null)
                                    //            _logger.WriteLog($"Closing Not Found For : {CustomScripName}", true);
                                    //        else
                                    //            TradePrice = ClosePrice.Close;
                                    //    }
                                    //    else
                                    //    {
                                    //        var ClosePrice = list_FOBhavcopy.Where(v => v.CustomScripname.Equals(CustomScripName)).FirstOrDefault();

                                    //        if (ClosePrice is null)
                                    //            _logger.WriteLog($"Closing Not Found For : {CustomScripName}", true);
                                    //        else if (UseClosing)
                                    //            TradePrice = ClosePrice.Close;
                                    //        else
                                    //            TradePrice = ClosePrice.SettlePrice;
                                    //    }
                                    //}
                                    //catch (Exception ee) { _logger.WriteLog("PS03 Loop Closing : " + line1 + Environment.NewLine + ee); }
                                    #endregion

                                    //added 20JAN2021 by Amey
                                    if (TradePrice <= 0)
                                    {
                                        _logger.Debug("Data Incorrect PS03 Skipped Invalid Qty/Price : " + Qty + "/" + TradePrice + Environment.NewLine + line1);
                                        continue;
                                    };

                                    //changed on 20APR2021 by Amey
                                    //changed on 12JAN2021 by Amey
                                    _EODPositionInfo = new EODPositionInfo()
                                    {
                                        Username = fields[7].Trim().ToUpper(),
                                        Segment = Segment,
                                        Token = ScripInfo.Token,
                                        TradePrice = TradePrice,
                                        TradeQuantity = Qty,
                                        UnderlyingSegment = ScripInfo.UnderlyingSegment,
                                        UnderlyingToken = ScripInfo.UnderlyingToken
                                    };

                                    list_Day1Positions.Add(_EODPositionInfo);
                                }
                                catch (Exception ee)
                                {
                                    _logger.Error(ee, "Data Incorrect PS03 : " + line1 + Environment.NewLine + ee);

                                    isAnyError = true;
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception Psex)
            {
                _logger.Debug("PS03 Upload " + Psex.ToString());

                isAnyError = true;
            }

            return list_Day1Positions;
        }

        //TODO: Make seperate class for such methods.
        private static double ConvertToUnixTimestamp(DateTime date)
        {
            DateTime origin = new DateTime(1980, 1, 1, 0, 0, 0, 0, DateTimeKind.Utc);
            TimeSpan diff = date - origin;
            return diff.TotalSeconds;
        }
    }
}
