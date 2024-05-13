using DevExpress.XtraEditors;
using DevExpress.XtraEditors.Filtering.Templates;
using Ionic.Zip;
using MySql.Data.MySqlClient;
using n.LicenseValidator;
using n.LicenseValidator.Data_Structures;
using n.Structs;
using NerveLog;
using NerveUtility;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using NSEUtilitaire;
using Org.BouncyCastle.Math;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Diagnostics.Contracts;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Security.Cryptography;
using System.Security.Policy;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Timers;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Linq;
using en_ScripType = NSEUtilitaire.en_ScripType;
using en_Segment = n.Structs.en_Segment;

namespace BOD_Utility
{
    public partial class Home : XtraForm
    {

        string ApplicationPath = Application.StartupPath + "\\";
        NerveLogger _logger = new NerveLogger(true, true, ApplicationName: "BOD-Utility");
        DataSet ds_Config = new DataSet();

        int SpanIndex = 0;
        int VaRIndex = 0;
        int BSESpanIndex = 0;
        int CDSpanIndex = 0;    //Addded by Akshay

        string[] arr_SpanFileExtensions;
        string[] arr_VaRFileExtensions;
        string[] arr_BSESpanFileExtensions;
        string[] arr_CDSpanFileExtensions;  //Added by Akshay

        Dictionary<string, FTPCRED> dict_FTPCred = new Dictionary<string, FTPCRED>();

        HashSet<string> hs_Usernames = new HashSet<string>();

        List<EODPositionInfo> list_Day1Positions = new List<EODPositionInfo>();

        /// <summary>
        /// Key : Segment|ScripName | Value : ScripInfo
        /// </summary>
        ConcurrentDictionary<string, ContractMaster> dict_ScripInfo = new ConcurrentDictionary<string, ContractMaster>();

        /// <summary>
        /// Key : Segment|CustomScripName | Value : ScripInfo
        /// </summary>
        ConcurrentDictionary<string, ContractMaster> dict_CustomScripInfo = new ConcurrentDictionary<string, ContractMaster>();

        /// <summary>
        /// Key : Segment|Token | Value : ScripInfo
        /// </summary>
        ConcurrentDictionary<string, ContractMaster> dict_TokenScripInfo = new ConcurrentDictionary<string, ContractMaster>();

        /// <summary>
        /// MySQL connection string.
        /// </summary>
        string _MySQLCon = string.Empty;

        bool UseUdiffFormat = false;

        /// <summary>
        /// Contents of Config.xml
        /// </summary>
        DataSet ds_SQLConfig = new DataSet();

        GatewayEngineConnector GatewayEngineConnector = new GatewayEngineConnector();

        bool IsWorking = true;

        List<string> list_ComponentStarted = new List<string>();

        Dictionary<string, string> dict_MappedClient = new Dictionary<string, string>();

        bool IsSpanFileDownloading = false;
        string SpanPath = string.Empty;
        bool SpanFileDownloaded = false;

        res_General res_APIResponse;

        string MemberCode = string.Empty;
        string APILoginID = string.Empty;
        string APIPassword = string.Empty;
        string SecretKey = string.Empty;

        int SpanWaitSeconds = 1000;

        bool isMcxContractFileContainsHeader = true;

        Config loadedConfig;//Added by Musharraf for new NSEUtilitaire dll 26-04-2023
        private XDocument xmlDoc;//Added by Musharraf April 10th 2023
        private LicenseInfo _LicenseInfo; //Added by Musharraf April 10th 2023
        #region upload file in db (Added by Musharraf 24th April 2023)
        private string CM_security_fileName;
        private string FO_contract_fileName;
        private string CD_contract_fileName;
        private string BSECM_security_fileName;
        private string FoSecban;
        private string DailySnapshot;
        private string MFHaircut;
        private string collateralHaircut;
        private string _NSE_CM_bhavcopy;
        private string _BSE_EQ_BHAVCOPY;
        private string _FOBhavcopy;
        private string _CDBhavcopy;
        private string _MCXbhavcopy;
        private string _MCXScripFile;
        #endregion
        public Home()
        {
            try
            {
                _logger.Initialize(ApplicationPath);

                //Added by Musharraf April 10th 2023 to read License
                var _licenseResponse = NerveLicenseValidator.Validate($@"{Application.StartupPath}\BOD.ns", "BOD", out /*licenseinfo*/ _LicenseInfo);
                _logger.Debug($" license response : {JsonConvert.SerializeObject(_LicenseInfo)}");

                //InitializeComponent();

                //_licenseinfo.enabledsegments.fo
                ///_licenseinfo.enabledsegments.fo
                if (_licenseResponse)
                {
                    InitializeComponent();
                    DateTime expirydate = _LicenseInfo.ExpiryDate;
                    
                    this.Text = string.Format("BOD Utility - License expires on {0:dd-MM-yyyy}", expirydate);//add expdate to the title bar
                }
                else
                {
                    _logger.Debug($"license error. {_LicenseInfo.Message} [{_LicenseInfo.Error}]");
                    XtraMessageBox.Show(_LicenseInfo.Message);
                    Environment.Exit(0);
                }
                //End of license reading logic


                /*InitializeComponent();*///Uncomment this when in Debugging and Comment the License one

                btn_Settings.Enabled = true;
                btn_StartAuto.Enabled = true;
                btn_RestartAuto.Enabled = true;


                ds_Config = NerveUtils.XMLC(ApplicationPath + "config.xml");//Old Config.xml
                xmlDoc = XDocument.Load(ApplicationPath + "config.xml");//To read new Config.xml nested tags Added by Musharraf April 10th 2023

                //added on 16MAR2021 by Amey
                var dRow = ds_Config.Tables["LOGIN"].Rows[0];
                var GuestCred = dRow["GUEST"].STR().SPL(',');
                dict_FTPCred.Add("GUEST", new FTPCRED() { Username = GuestCred[0], Password = GuestCred[1] });
                GuestCred = dRow["FO"].STR().SPL(',');
                dict_FTPCred.Add("FO", new FTPCRED() { Username = GuestCred[0], Password = GuestCred[1] });
                //Added by Akshay
                GuestCred = dRow["CD"].STR().SPL(',');
                dict_FTPCred.Add("CD", new FTPCRED() { Username = GuestCred[0], Password = GuestCred[1] });

                arr_SpanFileExtensions = ds_Config.GET("SAVEPATH", "SPAN-EXTENSTIONS").SPL(',');

                //Added by Akshay 
                arr_CDSpanFileExtensions = ds_Config.GET("SAVEPATH", "SPAN-EXTENSTIONS").SPL(',');

                //added on 05JAN2021 by Amey
                arr_VaRFileExtensions = ds_Config.GET("SAVEPATH", "VAR-EXTENSTIONS").SPL(',');

                //added on 05JAN2021 by Amey
                arr_BSESpanFileExtensions = ds_Config.GET("SAVEPATH", "BSE-SPAN-EXTENSTIONS").SPL(',');

                SpanWaitSeconds = Convert.ToInt32(ds_Config.GET("INTERVAL", "SPAN-WAIT-SECONDS")) * 1000;


                isMcxContractFileContainsHeader = Convert.ToBoolean(ds_Config.Tables["OTHER"].Rows[0]["MCX-CONTRACT-CONTAINS-HEADERS"].ToString());


                dateEdit_DownloadDate.DateTime = DateTime.Now;

                #region Added on 5-11-19 by Amey
                /// <summary>
                /// To Avoid => Exception : "The request was aborted: Could not create SSL/TLS secure channel"
                /// </summary>
                ServicePointManager.Expect100Continue = true;
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;

                #endregion

                //ReadConfig();
                //CheckMaxAllowedSqlPacket();
            }
            catch (Exception ee) { XtraMessageBox.Show("Error while initialising.", "Error"); _logger.Error(ee); }
        }

        HashSet<string> hs_ErrorIndex = new HashSet<string>();

        private void AddToList(string Message, bool IsError = false)
        {
            try
            {
                Message = $"{DateTime.Now} : {Message}";
                if (IsError)
                    hs_ErrorIndex.Add(Message);

                if (this.InvokeRequired)
                    this.Invoke((MethodInvoker)(() => { listBox_Messages.Items.Insert(0, Message); }));
                else
                    listBox_Messages.Items.Insert(0, Message);
            }
            catch (Exception ee) { _logger.Error(ee); }
        }

        
        async private void btn_StartMnually_Click(object sender, EventArgs e)
        {
            try
            {
                btn_StartMnually.Enabled = false;
                btn_Settings.Enabled = false;
                btn_StartAuto.Enabled = false;
                btn_RestartAuto.Enabled = false;
                btn_DownloadSpan.Enabled = false;

                _logger.Debug("Manual BOD Process Started, Event Triggered: btn_StartMnually_Click");

                //object tempObj = null;
                //ElapsedEventArgs tempE = null;
                //DownloadCDSpan(tempObj, tempE, new string[] { "C:\\Prime\\Other" });



                await Task.Run(() =>
                {

                  
                    //CHANGES ON 23DEC2022 BY NIKHIL | API DOWNLOAD
                    AddToList("Connecting to NSE API");

                    nNSEUtils.Instance.Initialize(ds_Config.GET("LOGIN", "MEMBER-CODE"), ds_Config.GET("LOGIN", "API-LOGINID"), ds_Config.GET("LOGIN", "API-PASSWORD"), ds_Config.GET("LOGIN", "SECRET-KEY"), Application.StartupPath + "\\config.json", out loadedConfig);
                    Console.WriteLine($"Config :  {JsonConvert.SerializeObject(loadedConfig)}");

                    res_APIResponse = nNSEUtils.Instance.LoginAPI(out res_LoginAPI _Response);
                    _logger.Debug("LOGIN REPONSE | STATUS : " + res_APIResponse.ResponseStatus + " | MESSAGE : " + res_APIResponse.Message + " | RESPONSE : " + res_APIResponse.Response.StatusDescription);
                    AddToList("API Connection Status | " + res_APIResponse.ResponseStatus);

                   
                    //Old Span files not needed and takes space in HDD. 22MAR2021 by Amey
                    //DeleteOldSpanDirectories();

                    DownloadDynamically();

                    /*

                    DownloadContractFile();//File name updated {uses API} 
                    Thread.Sleep(1000);

                    DownloadSecurityFile(); //File name updated {uses API}
                    Thread.Sleep(1000);

                    //Added by Akshay on 12-10-2021 for downloading CD contract
                    DownloadCDContractFile(); //File name updated {uses API}
                    Thread.Sleep(1000);

                    DownloadFOBhavcopy();//Modified by Musharraf 15th April 2023
                    Thread.Sleep(1000);
                    DownloadCMBhavcopy(((string)xmlDoc.Element("BOD-Utility").Element("CM").Element("BHAVCOPY").Element("WEBSITE")).Trim().Split(','), ds_Config.GET("SAVEPATH", "CM_BHAVCOPY").SPL(',')); //Modified by Musharraf 15th April 2023
                    Thread.Sleep(1000);

                    //Added by Akshay on 12 - 10 - 2021 for downloading CD Bhavcopy
                    DownloadCDBhavcopy(((string)xmlDoc.Element("BOD-Utility").Element("CD").Element("BHAVCOPY").Element("WEBSITE")).Trim().Split(','), ds_Config.GET("SAVEPATH", "CD_BHAVCOPY").SPL(','));//Modified by Musharraf 15th April 2023
                    Thread.Sleep(1000);

                    DownloadNNFSecurityFile();
                    Thread.Sleep(1000);
                    DownloadMFundHaircutFile();
                    Thread.Sleep(1000);
                    DownloadHaricutFile();
                    Thread.Sleep(1000);


                    DownloadFOSecBanFile(((string)xmlDoc.Element("BOD-Utility").Element("FO").Element("SECBAN").Element("WEBSITE")).Trim().Split(','), ds_Config.GET("SAVEPATH", "FO_SECBAN"));//Modified by Musharraf 17th April 2023
                    Thread.Sleep(1000);
                    DownloadSnapShot(((string)xmlDoc.Element("BOD-Utility").Element("FO").Element("DAILY_SNAPSHOT").Element("WEBSITE")).Trim().Split(','), ds_Config.GET("SAVEPATH", "DAILY_SNAPSHOT").SPL(','));//Modified by Musharraf 17th April 2023
                    Thread.Sleep(1000);
                    //added on 30APR2021 by Amey

                    DownloadBSEScripFile();//File name updated
                    Thread.Sleep(1000);
                    //added on 30APR2021 by Amey
                    DownloadBSECMBhavcopy(((string)xmlDoc.Element("BOD-Utility").Element("CM").Element("BSECM_BHAVCOPY").Element("WEBSITE")).Trim().Split(','), ds_Config.GET("SAVEPATH", "BSECM_BHAVCOPY").SPL(','));//Modified by Musharraf 17th April 2023
                    Thread.Sleep(1000);

                    DownloadMCXScrip();//Added by Musharraf for MCX Scrip File
                    Thread.Sleep(1000);
                    DownloadMCXBhavcopy(); //Added by Musharraf for MCXBhavcopy File
                    Thread.Sleep(1000);


                    InsertTokensIntoDBUdiff();
                    Thread.Sleep(5000);

                    var arr_SpanInfo = ds_Config.GET("SAVEPATH", "SPAN").SPL(',');
                    var arr_ExposureInfo = ds_Config.GET("SAVEPATH", "EXPOSURE").SPL(',');

                    for (int i = 0; i < arr_SpanInfo.Length; i++)
                    {
                        arr_SpanInfo[i] = arr_SpanInfo[i] + DateTime.Now.ToString("yyyyMMdd") + "\\";
                        arr_ExposureInfo[i] = arr_ExposureInfo[i] + DateTime.Now.ToString("yyyyMMdd") + "\\";

                        if (!Directory.Exists(arr_SpanInfo[i]))
                            Directory.CreateDirectory(arr_SpanInfo[i]);

                        if (!Directory.Exists(arr_ExposureInfo[i]))
                            Directory.CreateDirectory(arr_ExposureInfo[i]);
                    }

                    //added on 05JAN2021 by Amey
                    var VaRExposurePath = ds_Config.GET("SAVEPATH", "VAREXPOSURE");
                    if (!Directory.Exists(VaRExposurePath))
                        Directory.CreateDirectory(VaRExposurePath);

                    InvokeDownloader(arr_SpanInfo, arr_ExposureInfo, VaRExposurePath);
                    // ConvertPSO3Files() removed by musharraf
                    */
                });

                

                //btn_DownloadSpan.Enabled = true;
            }
            catch (Exception ee) { _logger.Error(ee, "Manual BOD Process: "); }

            

        }

        async private void btn_StartAuto_Click(object sender, EventArgs e)
        {
            try
            {
                btn_StartMnually.Enabled = false;
                btn_Settings.Enabled = false;
                btn_StartAuto.Enabled = false;
                btn_RestartAuto.Enabled = false;

                _logger.Debug("Automatic BOD Process Started, Event Triggered: btn_StartAuto_Click");

                await Task.Run(() =>
                {
                    AddToList("Connecting to NSE API");
                    nNSEUtils.Instance.Initialize(ds_Config.GET("LOGIN", "MEMBER-CODE"), ds_Config.GET("LOGIN", "API-LOGINID"), ds_Config.GET("LOGIN", "API-PASSWORD"), ds_Config.GET("LOGIN", "SECRET-KEY"), Application.StartupPath + "\\config.json", out loadedConfig);
                    res_APIResponse = nNSEUtils.Instance.LoginAPI(out res_LoginAPI _Response);
                    _logger.Debug("LOGIN REPONSE | STATUS : " + res_APIResponse.ResponseStatus + " | MESSAGE : " + res_APIResponse.Message);
                    AddToList("API Connection Status | " + res_APIResponse.ResponseStatus);

                    //Old Span files not needed and takes space in HDD. 22MAR2021 by Amey
                    //DeleteOldSpanDirectories();

                    DownloadDynamically();

                    /*

                    DownloadContractFile();
                    Thread.Sleep(1000);
                    DownloadSecurityFile();
                    Thread.Sleep(1000);

                    //Added by Akshay on 12-10-2021 for downloading CD contract
                    DownloadCDContractFile();
                    Thread.Sleep(1000);

                    //DownloadFOBhavcopy(ds_Config.GET("URLs", "FO_BHAVCOPY").SPL(','), ds_Config.GET("SAVEPATH", "FO_BHAVCOPY").SPL(','));
                    DownloadFOBhavcopy();
                    Thread.Sleep(1000);
                    //DownloadCMBhavcopy(ds_Config.GET("URLs", "CM_BHAVCOPY").SPL(','), ds_Config.GET("SAVEPATH", "CM_BHAVCOPY").SPL(','));
                    DownloadCMBhavcopy(((string)xmlDoc.Element("BOD-Utility").Element("CM").Element("BHAVCOPY").Element("WEBSITE")).Trim().Split(','), ds_Config.GET("SAVEPATH", "CM_BHAVCOPY").SPL(','));
                    Thread.Sleep(1000);

                    DownloadNNFSecurityFile();
                    Thread.Sleep(1000);
                    DownloadMFundHaircutFile();
                    Thread.Sleep(1000);
                    DownloadHaricutFile();
                    Thread.Sleep(1000);

                    //Added by Akshay on 12-10-2021 for downloading CD Bhavcopy
                    //DownloadCDBhavcopy(ds_Config.GET("URLs", "CD_BHAVCOPY").SPL(','), ds_Config.GET("SAVEPATH", "CD_BHAVCOPY").SPL(','));
                    DownloadCDBhavcopy(((string)xmlDoc.Element("BOD-Utility").Element("CD").Element("BHAVCOPY").Element("WEBSITE")).Trim().Split(','), ds_Config.GET("SAVEPATH", "CD_BHAVCOPY").SPL(','));
                    Thread.Sleep(1000);

                    //DownloadFOSecBanFile(ds_Config.GET("URLs", "FO_SECBAN").SPL(','), ds_Config.GET("SAVEPATH", "FO_SECBAN"));
                    DownloadFOSecBanFile(((string)xmlDoc.Element("BOD-Utility").Element("FO").Element("SECBAN").Element("WEBSITE")).Trim().Split(','), ds_Config.GET("SAVEPATH", "FO_SECBAN"));
                    Thread.Sleep(1000);
                    //DownloadSnapShot(ds_Config.GET("URLs", "DAILY_SNAPSHOT").SPL(','), ds_Config.GET("SAVEPATH", "DAILY_SNAPSHOT").SPL(','));
                    DownloadSnapShot(((string)xmlDoc.Element("BOD-Utility").Element("FO").Element("DAILY_SNAPSHOT").Element("WEBSITE")).Trim().Split(','), ds_Config.GET("SAVEPATH", "DAILY_SNAPSHOT").SPL(','));
                    Thread.Sleep(1000);
                    //added on 30APR2021 by Amey
                    DownloadBSEScripFile();
                    Thread.Sleep(1000);
                    //added on 30APR2021 by Amey
                    //DownloadBSECMBhavcopy(ds_Config.GET("URLs", "BSECM_BHAVCOPY").Split(','), ds_Config.GET("SAVEPATH", "BSECM_BHAVCOPY").SPL(','));
                    DownloadBSECMBhavcopy(((string)xmlDoc.Element("BOD-Utility").Element("CM").Element("BSECM_BHAVCOPY").Element("WEBSITE")).Trim().Split(','), ds_Config.GET("SAVEPATH", "BSECM_BHAVCOPY").SPL(','));//File name updated
                    Thread.Sleep(1000);

                    DownloadMCXScrip();//Added by Musharraf for MCX Scrip File
                    Thread.Sleep(1000);
                    DownloadMCXBhavcopy(); //Added by Musharraf for MCXBhavcopy File
                    Thread.Sleep(1000);

                    var arr_SpanInfo = ds_Config.GET("SAVEPATH", "SPAN").SPL(',');
                    var arr_ExposureInfo = ds_Config.GET("SAVEPATH", "EXPOSURE").SPL(',');

                    for (int i = 0; i < arr_SpanInfo.Length; i++)
                    {
                        arr_SpanInfo[i] = arr_SpanInfo[i] + DateTime.Now.ToString("yyyyMMdd") + "\\";
                        arr_ExposureInfo[i] = arr_ExposureInfo[i] + DateTime.Now.ToString("yyyyMMdd") + "\\";

                        if (!Directory.Exists(arr_SpanInfo[i]))
                            Directory.CreateDirectory(arr_SpanInfo[i]);

                        if (!Directory.Exists(arr_ExposureInfo[i]))
                            Directory.CreateDirectory(arr_ExposureInfo[i]);
                    }

                    //added on 05JAN2021 by Amey
                    var VaRExposurePath = ds_Config.GET("SAVEPATH", "VAREXPOSURE");
                    if (!Directory.Exists(VaRExposurePath))
                        Directory.CreateDirectory(VaRExposurePath);

                    InvokeDownloader(arr_SpanInfo, arr_ExposureInfo, VaRExposurePath);
                    //ConvertPSO3Files(); removed by musharraf

                    */
                });

                    
                StartComponents();     // Added by Snehadri on 15JUN2021 for Automatic BOD Process

            }
            catch (Exception ee) { _logger.Error(ee, "Automatic BOD Process: "); }
        }

        private void DownloadDynamically()
        {
            try
            { 

                //setting xml node for traversal of files through config.xml
                XmlDocument xmlDoc = new XmlDocument();
                XDocument xDoc = new XDocument();
                xDoc = XDocument.Load(ApplicationPath + "config.xml");
                xmlDoc.Load(ApplicationPath + "config.xml");
                XmlElement rootElm = xmlDoc.DocumentElement;
                XmlNode segmentSubnodesList = xmlDoc.SelectSingleNode("/BOD-Utility/SEGMENTS");


                //setting licenses 
                Dictionary<string, Boolean> enabledSegments = new Dictionary<string, Boolean>();
                enabledSegments.Add("FO", _LicenseInfo.EnabledSegments.FO);
                enabledSegments.Add("CD", _LicenseInfo.EnabledSegments.CD);
                enabledSegments.Add("MCX", _LicenseInfo.EnabledSegments.MCX);
                enabledSegments.Add("BSE", _LicenseInfo.EnabledSegments.BSE);
                enabledSegments.Add("CM", _LicenseInfo.EnabledSegments.CM);
                string segmentName;


                _logger.Debug("Traversing through segments");
                //Traversing across segments -- Chaitanya 03/05/2024
                foreach (XmlNode segment in segmentSubnodesList.ChildNodes)
                {
                    segmentName = segment.Name;
                    XmlNode fileTypes = xmlDoc.SelectSingleNode("/BOD-Utility/SEGMENTS/" + segmentName);



                    //Traversing over different filetypes one by one -- Chaitanya 03/05/2024
                    _logger.Debug("Traversing through " + segmentName);
                    foreach (XmlNode fileType in fileTypes.ChildNodes)
                    {
                        _logger.Debug("\n==========================================================================\n Downloading " + segmentName + " " + fileType.Name + " dynamically.\n ---------------------------------------------------------------");
                        AddToList(segmentName + " " + fileType.Name + " downloading........");
                        _logger.Debug($"Checking license status: EnbabledSegments." + segmentName + ": {_LicenseInfo.EnabledSegments." + segmentName + "}");
                        _logger.Debug("Fetching data from config.xml for file " + fileType.Name);

                        //Files sources and name
                        string website = fileType.SelectSingleNode("WEBSITE").InnerText;
                        string localpath = fileType.SelectSingleNode("LOCAL").InnerText;
                        var fileName = fileType.SelectSingleNode("NAME").InnerText;
                        string fileGenericName = fileType.Name; //eg. CONTRACT, SECURITY etc

                        //filename to search, to delete
                        string filenametodelete;

                        //logic to make filename searchable to delete
                        int index = fileGenericName.IndexOf("_0", StringComparison.OrdinalIgnoreCase);//removes substring starting from _0
                        filenametodelete = index != -1 ? fileGenericName.Substring(0, index) : fileGenericName;
                        int index2 = fileGenericName.IndexOf("_$", StringComparison.OrdinalIgnoreCase);//removes substring starting from _$
                        filenametodelete = index2 != -1 ? fileGenericName.Substring(0, index2) : fileGenericName;
                        int index3 = fileGenericName.IndexOf(".", StringComparison.OrdinalIgnoreCase);//removes substring starting from .
                        filenametodelete = index2 != -1 ? fileGenericName.Substring(0, index2) : fileGenericName;

                        //checking
                        if (!enabledSegments[segmentName])
                        {
                            return;
                        }
                        try
                        {

                            //savepath tag name
                            string savepathXmlTagName = segmentName + "_" + fileType.Name;
                            //getting savepath 
                            string[] arr_FolderPath = ((string) xDoc.Element("BOD-Utility").Element("SAVEPATH").Element(savepathXmlTagName)).Trim().Split(',');
                            string[] arr_OldFiles;



                            //logic to create directory(SAVEPATH) if not exists, and delete old files
                            bool directorycreated = false;
                            for (int i = 1; i < arr_FolderPath.Length; i++)
                            {
                                if (!Directory.Exists(arr_FolderPath[i]))
                                {
                                    Directory.CreateDirectory(arr_FolderPath[i]);
                                }
                                else
                                {
                                    arr_OldFiles = Directory.GetFiles(arr_FolderPath[i], filenametodelete);
                                    for (int j = 0; j < arr_OldFiles.Count(); j++)
                                    {
                                        File.Delete(arr_OldFiles[j]);
                                    }
                                }
                            }


                            var dateToCheck = dateEdit_DownloadDate.DateTime;
                            string respectiveWebsiteFileName = "";
                            //check for file in previous 7 days
                            Boolean filedownloaded = false;



                            for (int j = 0; j < 7; j++)
                            {
                                if (filedownloaded)
                                {
                                    break;
                                }

                                var filenameaccordingtoconfig = (fileName.Contains("$date:ddMMyyyy$") ? (fileName.Replace("$date:ddMMyyyy$", dateToCheck.STR("ddMMyyyy"))) : (fileName.Replace("$date:yyyyMMdd$", dateToCheck.STR("yyyyMMdd"))));


                                //filename for website
                                if (filenameaccordingtoconfig.EndsWith("csv"))
                                {
                                    respectiveWebsiteFileName = filenameaccordingtoconfig + ".zip";
                                }
                                else
                                {
                                    //respectiveWebsiteFileName = filenameaccordingtoconfig;
                                }

                                //incase file name is present within website eg CM BSE_SCRIP
                                if (website.EndsWith("zip"))
                                {
                                    respectiveWebsiteFileName = "";
                                }

                                //download using API
                                try
                                {
                                    //fetching apiUrls from config.json
                                    string binDebugPath = AppDomain.CurrentDomain.BaseDirectory;
                                    string jsonApplicationPath = Path.Combine(binDebugPath, "config.json");
                                    _logger.Debug("config path " + jsonApplicationPath);
                                    string jsonData = File.ReadAllText(jsonApplicationPath);

                                    JObject json = JObject.Parse(jsonData);
                                    string jsonBODfilepath = segmentName + "." + fileType.Name;

                                    string apiUrlPath = (string)json.SelectToken(jsonBODfilepath);
                                    _logger.Debug("api urlpath : " + apiUrlPath);

                                    //Downloading using API
                                    var response = nNSEUtils.Instance.DownloadCommonFile(segmentName, apiUrlPath, respectiveWebsiteFileName, arr_FolderPath[0]);
                                    //_logger.Debug($"Download"+segmentName+" "+fileType.Name+"API Response: " + JsonConvert.SerializeObject(response));



                                    if (response.ResponseStatus == en_ResponseStatus.SUCCESS)
                                    {
                                        filedownloaded = true;
                                        _logger.Debug($"after api success for {segmentName} {fileType.Name}------>" + filedownloaded);
                                        if (respectiveWebsiteFileName.EndsWith("zip"))
                                        {
                                            using (ZipFile zip = ZipFile.Read(arr_FolderPath[0] + respectiveWebsiteFileName))
                                            {
                                                zip.ExtractAll(arr_FolderPath[0], ExtractExistingFileAction.DoNotOverwrite);

                                                File.Delete(arr_FolderPath[0] + $"{respectiveWebsiteFileName}");

                                                //FOBhavcopyFilename = DecompressGZAndDelete(new FileInfo(arr_FolderPath[0] + respectiveFileName), string.Empty/*".csv"*/);


                                                AddToList(segmentName + " " + fileType + ": " + $"{filenameaccordingtoconfig} downloaded successfully using API ;)");
                                            }

                                            //Dont know what use
                                            /*_FOBhavcopy = BhavcopyFileName;
                                            break;*/
                                        }
                                        else if (respectiveWebsiteFileName.EndsWith("gz"))
                                        {
                                            FileInfo localFile = new FileInfo(arr_FolderPath[0] + respectiveWebsiteFileName);
                                            DecompressGZAndDelete(localFile, string.Empty/*".csv"*/);

                                        }
                                    }

                                }
                                catch (Exception ex)
                                {
                                    _logger.Error(ex, "Downloading " + segmentName + " " + fileType.Name + " from API: ");
                                }



                                //Lets download from WEBSITE
                                if (!filedownloaded)
                                {
                                    try
                                    {
                                        string url = website + respectiveWebsiteFileName;
                                        _logger.Debug("Fetched url for web >>>>>>>>>>>>>>>>" + url);
                                        //$"NSE_FO_bhavcopy_{dateToCheck.ToString("ddMMyyyy")}.csv";/*$"{dateEdit_DownloadDate.DateTime.STR("yyyy")}/{dateEdit_DownloadDate.DateTime.STR("MMM").UPP()}/fo{dateToCheck.STR("ddMMMyyyy").UPP()}bhav.csv.zip";*/
                                        //BhavcopyFileName = $"NSE_FO_bhavcopy_{dateToCheck.ToString("ddMMyyyy")}.csv";
                                        using (WebClient webClient = new WebClient())
                                        {
                                            _logger.Debug("Savepath is " + arr_FolderPath[0] + filenameaccordingtoconfig);
                                            webClient.DownloadFile(url, arr_FolderPath[0] + respectiveWebsiteFileName);
                                            filedownloaded = true;
                                            _logger.Debug($"after website success for {segmentName} {fileType.Name}------>" + filedownloaded);

                                        }
                                        if (respectiveWebsiteFileName.EndsWith(".zip"))
                                        {
                                            using (ZipFile zip = ZipFile.Read(arr_FolderPath[0] + respectiveWebsiteFileName))
                                            {
                                                zip.ExtractAll(arr_FolderPath[0], ExtractExistingFileAction.DoNotOverwrite);
                                                File.Delete(arr_FolderPath[0] + $"{respectiveWebsiteFileName}");
                                            }
                                        }
                                        else if (respectiveWebsiteFileName.EndsWith("gz"))
                                        {
                                            FileInfo localFile = new FileInfo(arr_FolderPath[0] + respectiveWebsiteFileName);
                                            DecompressGZAndDelete(localFile, string.Empty/*".csv"*/);
                                        }

                                        //File.Delete(arr_BhavcopyFolderPath[0] + BhavcopyFileName.Replace(".gz", ".zip"));
                                        AddToList(segmentName + " " + fileType.Name + $" {respectiveWebsiteFileName} downloaded successfully using WEBSITE =)");
                                        //_FOBhavcopy = BhavcopyFileName; //BhavcopyFileName.Substring(0, BhavcopyFileName.LastIndexOf(".csv") + 4);
                                        break;
                                    }
                                    catch (Exception ee)
                                    {
                                        _logger.Error(ee, "Downloading" + segmentName + " " + fileType.Name + " from Website:");
                                        _logger.Debug($"Url Passed: {website + respectiveWebsiteFileName}");
                                    }
                                }


                                _logger.Debug($"after website success for 2 {segmentName} {fileType.Name}------>" + filedownloaded);
                                //lets download from localpath
                                if (!filedownloaded)
                                {
                                    _logger.Debug("donwload from localpath not gettinng execute now");
                                    try
                                    {
                                        File.Copy(localpath + filenameaccordingtoconfig, arr_FolderPath[0] + filenameaccordingtoconfig, true);
                                        AddToList(segmentName + " " + fileType.Name + ": " + $"{filenameaccordingtoconfig} copied successfully from localpath.");

                                        if (filenameaccordingtoconfig.EndsWith(".zip"))
                                        {
                                            using (ZipFile zip = ZipFile.Read(arr_FolderPath[0] + filenameaccordingtoconfig))
                                            {
                                                zip.ExtractAll(arr_FolderPath[0], ExtractExistingFileAction.DoNotOverwrite);
                                            }
                                        }
                                        else if (filenameaccordingtoconfig.EndsWith("gz"))
                                        {
                                            FileInfo localFile = new FileInfo(arr_FolderPath[0] + filenameaccordingtoconfig);
                                            DecompressGZAndDelete(localFile, string.Empty/*".csv"*/);
                                        }
                                        //_FOBhavcopy = BhavcopyFileName;
                                        filedownloaded = true;
                                        break;
                                    }
                                    catch (Exception ee)
                                    {
                                        _logger.Error(ee, "Copying " + segmentName + " " + fileType.Name + " from Local Folder:");
                                        _logger.Debug($"Source : {localpath + filenameaccordingtoconfig} and Destination: {arr_FolderPath[0] + filenameaccordingtoconfig}");
                                    }
                                }
                                _logger.Debug($"after localpath success for {segmentName} {fileType.Name}------>" + filedownloaded);

                                // subtract a day from the date to check the previous day
                                dateToCheck = dateToCheck.AddDays(-1);

                                // skip weekends
                                if (dateToCheck.DayOfWeek == DayOfWeek.Saturday)
                                {
                                    dateToCheck = dateToCheck.AddDays(-1);
                                }
                                else if (dateToCheck.DayOfWeek == DayOfWeek.Sunday)
                                {
                                    dateToCheck = dateToCheck.AddDays(-2);
                                }


                            }



                            if (!filedownloaded)
                            {
                                _logger.Debug(">>>>>>>>>>>>Download failed for " + segmentName + " " + respectiveWebsiteFileName + "\n ========================================================================");
                                AddToList("Download failed for " + segmentName + " " + fileName);
                            }

                        }


                        catch (Exception ee)
                        {
                            _logger.Error(ee, "Download " + segmentName + " " + fileType.Name + $" with {website}");
                        }

                    }

                }
            }
            catch (Exception ee)
            {
                _logger.Error(ee, "DynamicDownloading");
            }
        }

        private void DeleteOldSpanDirectories()
        {
            _logger.Debug("Executing DeleteOldSpanDirectories(): ");
            try
            {
                var arr_SpanInfo = ds_Config.GET("SAVEPATH", "SPAN").SPL(',');

                foreach (var _SpnaDir in arr_SpanInfo)
                {
                    string[] subdirs = Directory.GetDirectories(_SpnaDir);

                    foreach (var item in subdirs)
                    {
                        try
                        {
                            var FolderDate = DateTime.ParseExact(item.SUB(item.LastIndexOf('\\') + 1), "yyyyMMdd", CultureInfo.InvariantCulture);
                            if (FolderDate.Date < DateTime.Now.AddDays(-5))
                                Directory.Delete(item, true);
                        }
                        catch (Exception ee) { _logger.Error(ee, "DeleteOldSpanDirectories Loop : " + item); }
                    }
                }


                var arr_SpanLogPath = ds_Config.GET("SAVEPATH", "SPAN-LOGS").SPL(',');

                foreach (var _SpanLogPath in arr_SpanLogPath)
                {
                    foreach (var item in Directory.GetFiles(_SpanLogPath))
                    {
                        try
                        {
                            var _ExtractedDate = item.SUB(item.LastIndexOf('\\') + 1).SPL('_')[1].SPL('.')[0];
                            var FolderDate = DateTime.ParseExact(_ExtractedDate, "yyyyMMdd", CultureInfo.InvariantCulture);
                            if (FolderDate.Date < DateTime.Now.AddDays(-5))
                                File.Delete(item);
                        }
                        catch (Exception ee) { _logger.Error(ee, "DeleteOldSpanDirectories -CONTRACT Loop : " + item); }
                    }
                }

                var arr_VaRExposurepath = ds_Config.GET("SAVEPATH", "VAREXPOSURE");

                foreach (var item in Directory.GetFiles(arr_VaRExposurepath))
                {
                    try
                    {
                        var _ExtractedDate = item.SUB(item.LastIndexOf('\\') + 1).SPL('_')[2];
                        var FolderDate = DateTime.ParseExact(_ExtractedDate, "ddMMyyyy", CultureInfo.InvariantCulture);
                        if (FolderDate.Date < DateTime.Now.AddDays(-5))
                            File.Delete(item);
                    }
                    catch (Exception ee) { _logger.Error(ee, "DeleteOldSpanDirectories -VaRExosure Loop : " + item); }
                }
            }
            catch (Exception ee) { _logger.Error(ee, "DeleteOldSpanDirectories()"); }
        }

        List<DayOfWeek> lst_WeekendDays = new List<DayOfWeek>();
        

        private void DownloadContractFile()
        {
            
            _logger.Debug($"Checking license status: EnabledSegments.FO: {_LicenseInfo.EnabledSegments.FO}");//Added by Musharraf 10th April 2023
                                                                                                             //
            if (!_LicenseInfo.EnabledSegments.FO)
            {
                return;
            }

            try
            {
                _logger.Debug("Executing  DownloadContractFile(): ");

                string[] arr_ContractFolderPath = ds_Config.GET("SAVEPATH", "CONTRACT").SPL(',');
                //string[] arr_ContractURL = ds_Config.GET("URLs", "CONTRACT").SPL(',');
                string downloadFTP = ((string)xmlDoc.Element("BOD-Utility").Element("FO").Element("CONTRACT").Element("FTP")).Trim();//To download file FTP
                string[] arr_ContractURL = downloadFTP.Split(',');

                string downloadWebsite = ((string)xmlDoc.Element("BOD-Utility").Element("FO").Element("CONTRACT").Element("WEBSITE")).Trim();//To download Website
                string[] arr_ContractWebURL = downloadWebsite.Split(',');

                string downloadFromLocal = ((string)xmlDoc.Element("BOD-Utility").Element("FO").Element("CONTRACT").Element("LOCAL")).Trim();//To download from local file
                string[] arr_ContractLocalFile = downloadFromLocal.Split(',');

                var FileName = ((string)xmlDoc.Element("BOD-Utility").Element("FO").Element("CONTRACT").Element("NAME")).Trim();//FileName of your Contract file

                bool filedownloaded = false;

                if (arr_ContractFolderPath.Length != 0)
                {
                    AddToList("Contract file downloading.");

                    for (int i = 0; i < arr_ContractFolderPath.Length; i++)
                    {
                        if (!Directory.Exists(arr_ContractFolderPath[i]))
                            Directory.CreateDirectory(arr_ContractFolderPath[i]);
                        else
                        {
                            string[] files = Directory.GetFiles(arr_ContractFolderPath[i], @"NSE_FO_contract_*.csv");//Changed from *.gz to "NSE_FO_contract_*.csv" 
                            for (int j = 0; j < files.Count(); j++)
                                File.Delete(files[j]);
                        }
                    }

                    var dateToCheck = dateEdit_DownloadDate.DateTime;

                    //  Added by Musharraf 3rd April 2023    Latest file from the previous 7 working day's 
                    // check if today is a weekend day
                    //if (dateToCheck.DayOfWeek == DayOfWeek.Saturday || dateToCheck.DayOfWeek == DayOfWeek.Sunday)
                    //{
                    //    // if so, set the date to the previous Friday
                    //    dateToCheck = (dateToCheck.DayOfWeek == DayOfWeek.Saturday) ? dateToCheck.AddDays(-1) : dateToCheck.AddDays(-2);
                    //}

                    for (int i = 0; i < 7; i++)
                    {
                        // check for the file in the previous 7 working days
                        try
                        {
                            res_General file_check = null;
                            //var NSEFOContractFile = ds_Config.GET("FILENAME", "FO_CONTRACT").replace("$date:ddMMyyyy$", dateToCheck.ToString("ddMMyyyy"));//"NSE_FO_contract_" + DateTime.Now.ToString("ddMMyyyy") + ".csv"; //Added by Musharraf
                            //Added by Musharraf to check previous 7 days files
                            var NSEFOContractFile = FileName.Replace("$date:ddMMyyyy$", dateToCheck.ToString("ddMMyyyy"));//FileName replaced this ds_Config.GET("FILENAME", "FO_CONTRACT").ToString()
                                                                                                                          //if (latestTxtFile == null || latestTxtFile.LastWriteTime.Date != DateTime.Today){
                                                                                                                          //downloading with NSE-API
                                                                                                                          //NSEFOContractFile = "NSE_FO_contract_" + dateToCheck.ToString("ddMMyyyy") + ".csv";
                            file_check = nNSEUtils.Instance.DownloadCommonFile(en_FolderTypes.FO_CONTRACT, NSEFOContractFile, arr_ContractFolderPath[0]);
                            _logger.Debug($"DownloadContractFile API Response: " + JsonConvert.SerializeObject(file_check));
                            //downloadTask.Wait();
                            // download the file if it exists
                            if (file_check != null && file_check.ResponseStatus == en_ResponseStatus.SUCCESS)
                            {
                                //_logger.Debug("DownloadContractFile API Response: " + JsonConvert.SerializeObject(file_check.Response.ResponseStatus));

                                if (NSEFOContractFile.EndsWith(".gz") && File.Exists(Path.Combine(arr_ContractFolderPath[0], NSEFOContractFile)))
                                {
                                    _logger.Debug("inside If NSEFOContractFile.EndsWith(.gz) && File.Exists(Path.Combine(arr_ContractFolderPath[0], NSEFOContractFile))");
                                    FileInfo localFile = new FileInfo(arr_ContractFolderPath[0] + NSEFOContractFile);
                                    NSEFOContractFile = DecompressGZAndDelete(localFile, "");
                                    NSEFOContractFile = Path.GetFileName(NSEFOContractFile);
                                }
                                
                                filedownloaded = true;
                                AddToList($"Contract file: {NSEFOContractFile} downloaded successfully.");
                                _logger.Debug($"Downloaded Contract File From API : {NSEFOContractFile}");
                            }


                            if (filedownloaded)
                            {
                                // exit the loop if the file is downloaded successfully
                                FO_contract_fileName = NSEFOContractFile;
                                break;
                            }

                        }
                        catch (Exception ee)
                        {
                            _logger.Error(ee, "Downloading Contract File From API :");
                        }

                        #region FTP Download(No longer used)
                        //if (!filedownloaded && !string.IsNullOrEmpty(arr_ContractURL[0]))
                        //{
                        //    try
                        //    {
                        //        using (WebClient webClient = new WebClient())
                        //        {
                        //            //Added to login and download from NSE FTP link. 16MAR2021-Amey
                        //            webClient.Credentials = new NetworkCredential(dict_FTPCred["FO"].Username, dict_FTPCred["FO"].Password);

                        //            webClient.DownloadFile(arr_ContractURL[0], arr_ContractFolderPath[0] + @"contract.gz");
                        //        }

                        //        File.Delete(arr_ContractFolderPath[0] + @"NSE_FO_contract_" + dateEdit_DownloadDate.DateTime.ToString("ddMMyyyy") + ".csv");

                        //        DecompressGZAndDelete(new FileInfo(arr_ContractFolderPath[0] + @"contract.gz"), ".txt");

                        //        filedownloaded = true;
                        //        if (filedownloaded == true)
                        //        {
                        //            break;
                        //        }
                        //    }
                        //    catch (Exception ee) { _logger.Error(ee, "DownloadExchangeFiles - Contract"); }
                        //}
                        #endregion

                        try
                        {
                            var ContractFile = FileName.Replace("$date:ddMMyyyy$", dateToCheck.ToString("ddMMyyyy"));

                            if (!filedownloaded && !string.IsNullOrEmpty(arr_ContractLocalFile[0]))
                            {
                                if (ContractFile.EndsWith(".gz") && File.Exists(Path.Combine(arr_ContractLocalFile[0], ContractFile)))
                                {
                                    FileInfo localFile = new FileInfo(arr_ContractLocalFile[0] + ContractFile);
                                    ContractFile = DecompressGZAndDelete(localFile, "");
                                    ContractFile = Path.GetFileName(ContractFile);
                                }
                                else
                                {
                                    ContractFile = ContractFile.Replace(".gz", "");
                                }

                                string source = arr_ContractLocalFile[0] + ContractFile;
                                string destination = arr_ContractFolderPath[0] + ContractFile;
                                File.Copy(source, destination, true);
                                filedownloaded = true;
                                AddToList($"Contract file: {ContractFile} downloaded successfully.");
                                if (filedownloaded == true)
                                {
                                    FO_contract_fileName = ContractFile;
                                    break;
                                }
                            }
                        }
                        catch (Exception ee) { _logger.Error(ee, "Copying Contract File From Local Folder:"); }

                        // subtract a day from the date to check the previous day
                        dateToCheck = dateToCheck.AddDays(-1);

                        //// skip weekends
                        if (dateToCheck.DayOfWeek == DayOfWeek.Saturday)
                        {
                            dateToCheck = dateToCheck.AddDays(-1);
                        }
                        else if (dateToCheck.DayOfWeek == DayOfWeek.Sunday)
                        {
                            dateToCheck = dateToCheck.AddDays(-2);
                        }
                    }

                    //end of the previous 7 days file check
                    if (!filedownloaded)
                    {
                        AddToList("Contract file download failed.", true);
                    }

                    if (arr_ContractFolderPath.Length > 1 && filedownloaded)
                    {
                        for (int i = 1; i < arr_ContractFolderPath.Length; i++)
                            File.Copy(arr_ContractFolderPath[0] + FO_contract_fileName, arr_ContractFolderPath[i] + FO_contract_fileName, true);
                        _logger.Debug("Contract file downloaded successfully in all Save-Paths.");
                    }
                }
                else
                    AddToList("Invalid path specified for Contract file.", true);
            }
            catch (Exception ee) { _logger.Error(ee); AddToList("Unable to download Contract file.", true); }
        }
        private void DownloadSecurityFile()
        {
            _logger.Debug($"Checking license status: EnabledSegments.CM: {_LicenseInfo.EnabledSegments.CM}");// Added by Musharraf 10th April 2023
            if (!_LicenseInfo.EnabledSegments.CM)
            {
                return;
            }
            try
            {
                _logger.Debug("Executing DownloadSecurityFile()");

                string[] arr_SecurityFolderPath = ds_Config.GET("SAVEPATH", "SECURITY").SPL(',');
                //string[] arr_SecurityUrl = ds_Config.GET("URLs", "SECURITY").SPL(',');

                string downloadFTP = ((string)xmlDoc.Element("BOD-Utility").Element("CM").Element("SECURITY").Element("FTP")).Trim();//To download file FTP
                string[] arr_SecurityUrl = downloadFTP.Split(',');

                string downloadWebsite = ((string)xmlDoc.Element("BOD-Utility").Element("CM").Element("SECURITY").Element("WEBSITE")).Trim();//To download Website
                string[] arr_ContractWebURL = downloadWebsite.Split(',');

                string downloadFromLocal = ((string)xmlDoc.Element("BOD-Utility").Element("CM").Element("SECURITY").Element("LOCAL")).Trim();//To download from local file
                string[] arr_ContractLocalFile = downloadFromLocal.Split(',');

                var FileName = ((string)xmlDoc.Element("BOD-Utility").Element("CM").Element("SECURITY").Element("NAME")).Trim();//FileName of your Contract file

                bool filedownloaded = false;

                if (arr_SecurityFolderPath.Length != 0)
                {
                    AddToList("Security file downloading.");

                    for (int i = 0; i < arr_SecurityFolderPath.Length; i++)
                    {
                        if (!Directory.Exists(arr_SecurityFolderPath[i]))
                            Directory.CreateDirectory(arr_SecurityFolderPath[i]);
                        else
                        {
                            string[] files = Directory.GetFiles(arr_SecurityFolderPath[i], @"NSE_CM_security_*.csv");
                            for (int j = 0; j < files.Count(); j++)
                                File.Delete(files[j]);
                        }
                    }

                    var dateToCheck = dateEdit_DownloadDate.DateTime;
                    var NSECMSecurity = FileName.Replace("$date:ddMMyyyy$", dateToCheck.ToString("ddMMyyyy"));

                    // check if today is a weekend day
                    //if (dateToCheck.DayOfWeek == DayOfWeek.Saturday || dateToCheck.DayOfWeek == DayOfWeek.Sunday)
                    //{
                    //    // if so, set the date to the previous Friday
                    //    dateToCheck = (dateToCheck.DayOfWeek == DayOfWeek.Saturday) ? dateToCheck.AddDays(-1) : dateToCheck.AddDays(-2);
                    //}
                    // Added by Musharraf 3rd April 2023
                    for (int i = 0; i < 7; i++)
                    {
                        //Added by Musharraf to check previous 7 days files
                        try
                        {
                            res_General file_check = null;
                            // check for the file in the previous 7 working days
                            NSECMSecurity = FileName.Replace("$date:ddMMyyyy$", dateToCheck.ToString("ddMMyyyy"));
                            //NSECMSecurity = "NSE_CM_security_" + dateToCheck.ToString("ddMMyyyy") + ".csv";
                            file_check = nNSEUtils.Instance.DownloadCommonFile(en_FolderTypes.CM_SECURITY, NSECMSecurity, arr_SecurityFolderPath[0]);
                            _logger.Debug("DownloadSecurityFile API Response: " + JsonConvert.SerializeObject(file_check));
                            //Task<res_General> downloadTask = Task.Run(() =>
                            //{
                            //    return file_check=nNSEUtils.Instance.DownloadCommonFile(en_FolderTypes.CM_SECURITY, NSECMSecurity, arr_SecurityFolderPath[0]);
                            //});

                            // download the file if it exists  
                            if (file_check != null && file_check.ResponseStatus == en_ResponseStatus.SUCCESS)
                            {
                                // exit the loop if the file is downloaded successfully
                                //_logger.Debug("DownloadContractFile API Response: " + JsonConvert.SerializeObject(file_check.Response.ResponseStatus));
                                if (NSECMSecurity.EndsWith(".gz") && File.Exists(Path.Combine(arr_SecurityFolderPath[0], NSECMSecurity)))
                                {
                                    _logger.Debug("Inside if NSECMSecurity.EndsWith(.gz) && File.Exists(Path.Combine(arr_SecurityFolderPath[0], NSECMSecurity))");
                                    FileInfo localFile = new FileInfo(arr_SecurityFolderPath[0] + NSECMSecurity);
                                    NSECMSecurity = DecompressGZAndDelete(localFile, "");
                                    NSECMSecurity = Path.GetFileName(NSECMSecurity);
                                }
                                
                                filedownloaded = true;
                                AddToList($"Security file: {NSECMSecurity} downloaded successfully.");
                                _logger.Debug($"Downloaded Security File From API : {NSECMSecurity}");
                            }


                            if (filedownloaded)
                            {
                                CM_security_fileName = NSECMSecurity;
                                break;
                            }
                            //var response = nNSEUtils.Instance.DownloadCommonFile(en_FolderTypes.CM_SECURITY, NSECMSecurity/*"security.gz"*/, arr_SecurityFolderPath[0]);
                            //_logger.Debug("DownloadSecurityFile API Response: " + JsonConvert.SerializeObject(response));

                        }
                        catch (Exception ee)
                        {
                            _logger.Error(ee, "Downloading Security File From API ");
                        }
                        #region FTP download(Decommissioned)
                        //if (!filedownloaded && !string.IsNullOrEmpty(arr_SecurityUrl[0]))
                        //{

                        //    try
                        //    {
                        //        using (WebClient webClient = new WebClient())
                        //        {
                        //            //Added to login and download from NSE FTP link. 16MAR2021-Amey
                        //            webClient.Credentials = new NetworkCredential(dict_FTPCred["FO"].Username, dict_FTPCred["FO"].Password);

                        //            webClient.DownloadFile(arr_SecurityUrl[0], arr_SecurityFolderPath[0] + @"security.gz");
                        //        }

                        //        File.Delete(arr_SecurityFolderPath[0] + @"NSE_CM_security_" + DateTime.Now.ToString("ddMMyyyy") + ".csv");

                        //        DecompressGZAndDelete(new FileInfo(arr_SecurityFolderPath[0] + @"security.gz"), ".txt");
                        //        filedownloaded = true;
                        //    }
                        //    catch (Exception ee) { _logger.Error(ee, "DownloadExchangeFiles - Security"); }
                        //}
                        #endregion
                        try
                        {
                            NSECMSecurity = FileName.Replace("$date:ddMMyyyy$", dateToCheck.ToString("ddMMyyyy"));

                            if (!filedownloaded && !string.IsNullOrEmpty(arr_ContractLocalFile[0]))
                            {
                                if (NSECMSecurity.EndsWith(".gz") && File.Exists(Path.Combine(arr_ContractLocalFile[0], NSECMSecurity)))
                                {
                                    FileInfo localFile = new FileInfo(arr_ContractLocalFile[0] + NSECMSecurity);
                                    NSECMSecurity = DecompressGZAndDelete(localFile, "");
                                    NSECMSecurity = Path.GetFileName(NSECMSecurity);
                                }
                                else
                                {
                                    NSECMSecurity = NSECMSecurity.Replace(".gz", "");
                                }

                                string source = arr_ContractLocalFile[0] + NSECMSecurity;
                                string destination = arr_SecurityFolderPath[0] + NSECMSecurity;
                                File.Copy(source, destination, true);
                                filedownloaded = true;
                                AddToList($"Security file: {NSECMSecurity} downloaded successfully.");
                                if (filedownloaded == true)
                                {
                                    CM_security_fileName = NSECMSecurity;
                                    break;
                                }
                            }
                        }
                        catch (Exception ee) { _logger.Error(ee, "Copying Security File From Local Folder:"); }

                        // subtract a day from the date to check the previous day
                        dateToCheck = dateToCheck.AddDays(-1);

                        // skip weekends
                        if (dateToCheck.DayOfWeek == DayOfWeek.Saturday)
                        {
                            dateToCheck = dateToCheck.AddDays(-1);
                        }
                        else if (dateToCheck.DayOfWeek == DayOfWeek.Sunday)
                        {
                            dateToCheck = dateToCheck.AddDays(-2);
                        }
                    }

                    //end of the previous 7 days file check

                    if (!filedownloaded)
                    {
                        AddToList("Security file download failed.", true);
                        return;
                    }

                    if (arr_SecurityFolderPath.Length > 1 && filedownloaded)
                    {
                        for (int i = 1; i < arr_SecurityFolderPath.Length; i++)
                            File.Copy(arr_SecurityFolderPath[0] + CM_security_fileName, arr_SecurityFolderPath[i] + CM_security_fileName, true);

                        _logger.Debug("Security file downloaded successfully in all Save-Paths");
                    }
                }
                else
                    AddToList("Invalid path specified for Security file.", true);
            }
            catch (Exception ee) { _logger.Error(ee); AddToList("Unable to download Security file.", true); }
        }


        private void DownloadNNFSecurityFile()
        {
            _logger.Debug($"Checking license status: EnabledSegments.CM: {_LicenseInfo.EnabledSegments.CM}"); // Added by Musharraf 3rd April 2023
            if (!_LicenseInfo.EnabledSegments.CM)
            {
                return;
            }
            try
            {
                _logger.Debug("Executing  DownloadNNFSecurityFile():");

                string[] arr_NNFSecurityFolderPath = ds_Config.GET("SAVEPATH", "NNF-SECURITY").SPL(',');
                //string[] arr_NNFSecurityURL = ds_Config.GET("URLs", "NNF-SECURITY").SPL(',');
                var nnfFileName = ((string)xmlDoc.Element("BOD-Utility").Element("CM").Element("NNF-SECURITY").Element("NAME")).Trim();//FileName of your file
                string[] arr_LocalFolder = ((string)xmlDoc.Element("BOD-Utility").Element("CM").Element("NNF-SECURITY").Element("LOCAL")).Trim().Split();
                bool filedownloaded = false;

                if (arr_NNFSecurityFolderPath.Length != 0)
                {
                    AddToList("NNF Security file downloading.");

                    for (int i = 0; i < arr_NNFSecurityFolderPath.Length; i++)
                    {
                        if (!Directory.Exists(arr_NNFSecurityFolderPath[i]))
                            Directory.CreateDirectory(arr_NNFSecurityFolderPath[i]);
                        else
                        {
                            string[] files = Directory.GetFiles(arr_NNFSecurityFolderPath[i], @".gz");
                            for (int j = 0; j < files.Count(); j++)
                                File.Delete(files[j]);
                        }
                    }

                    try
                    {
                        var response = nNSEUtils.Instance.DownloadCommonFile(en_FolderTypes.CM_NNF_SECUIRTY, nnfFileName, arr_NNFSecurityFolderPath[0]);
                        _logger.Debug("DownloadSecurityFile API Response: " + JsonConvert.SerializeObject(response));
                        if (response.ResponseStatus == en_ResponseStatus.SUCCESS)
                        {
                            DecompressGZAndDelete(new FileInfo(arr_NNFSecurityFolderPath[0] + @"nnf_security.gz"), ".txt");
                            filedownloaded = true;
                            AddToList($"NNF Security file: {nnfFileName.Replace(".gz", ".txt")} downloaded successfully.");
                        }
                    }
                    catch (Exception ee)
                    {
                        _logger.Error(ee, "Downloading NNF Security File From API :");
                    }
                    #region FTP
                    //if (!filedownloaded)
                    //{
                    //    try
                    //    {
                    //        using (WebClient webClient = new WebClient())
                    //        {
                    //            //Added to login and download from NSE FTP link. 16MAR2021-Amey
                    //            webClient.Credentials = new NetworkCredential(dict_FTPCred["FO"].Username, dict_FTPCred["FO"].Password);
                    //            webClient.DownloadFile(arr_NNFSecurityURL[0], arr_NNFSecurityFolderPath[0] + @"nnf_security.gz");
                    //        }

                    //        File.Delete(arr_NNFSecurityFolderPath[0] + @"nnf_security.dat");

                    //        DecompressGZAndDelete(new FileInfo(arr_NNFSecurityFolderPath[0] + @"nnf_security.gz"), ".txt");

                    //        filedownloaded = true;
                    //    }
                    //    catch (Exception ee)
                    //    {
                    //        _logger.Error(ee, "DownloadExchangeFiles - NNF-SECURITY");
                    //    }
                    //}
                    #endregion
                    try
                    {
                        //Local folder
                        if (!filedownloaded && !string.IsNullOrEmpty(arr_LocalFolder[0]))
                        {
                            var nnfFile = "nnf_security.txt";
                            File.Copy(arr_LocalFolder[0] + nnfFile, arr_NNFSecurityFolderPath[0] + nnfFile, true);

                            filedownloaded = true;

                            AddToList($"NNF Security file: {nnfFile} downloaded successfully.");
                        }
                    }
                    catch (Exception ee)
                    {
                        _logger.Error(ee, $"Copying NNF Security file From Local Folder:");
                        _logger.Debug($"{nnfFileName} missing from path {arr_LocalFolder}");
                        throw;
                    }
                    if (!filedownloaded)
                    {
                        AddToList("NNF Security file download failed.", true);
                        return;
                    }

                    if (arr_NNFSecurityFolderPath.Length > 1 && filedownloaded)
                    {
                        // AddToList("NNF Security file downloaded successfully.");
                        for (int i = 1; i < arr_NNFSecurityFolderPath.Length; i++)
                            File.Copy(arr_NNFSecurityFolderPath[0] + @"nnf_security.txt", arr_NNFSecurityFolderPath[i] + @"nnf_security.txt", true);
                    }
                    if (filedownloaded) { _logger.Debug("NNF Security file downloaded successfully in all Save-Paths."); }
                }
                else
                    AddToList("Invalid path specified for NNF Security file.", true);
            }
            catch (Exception ee)
            {
                _logger.Error(ee); AddToList("Unable to download NNF Security file.", true);
            }
        }

        private void DownloadMFundHaircutFile()
        {
            _logger.Debug($"Checking license status: EnabledSegments.CM and FO: {_LicenseInfo.EnabledSegments.CM} and {_LicenseInfo.EnabledSegments.FO}"); // Added by Musharraf 3rd April 2023
            if (!(_LicenseInfo.EnabledSegments.CM && _LicenseInfo.EnabledSegments.FO))
            {
                return;
            }
            try
            {
                _logger.Debug("Executing DownloadMFundHaircutFile(): ");

                string[] arr_MFundHaircutFolderPath = ds_Config.GET("SAVEPATH", "MF-HAIRCUT").SPL(',');
                //string[] arr_MFundHaircutURL = ds_Config.GET("URLs", "MF-HAIRCUT").SPL(',');
                string downloadWebsite = ((string)xmlDoc.Element("BOD-Utility").Element("FO").Element("MF-HAIRCUT").Element("WEBSITE")).Trim();//To download Website
                string[] arr_MFundHaircutURL = downloadWebsite.Split(',');

                string Local = ((string)xmlDoc.Element("BOD-Utility").Element("FO").Element("MF-HAIRCUT").Element("LOCAL")).Trim();//To download Website
                string[] arr_MFundHaircutURLlocal = Local.Split(',');

                string HaircutFileName = ((string)xmlDoc.Element("BOD-Utility").Element("FO").Element("MF-HAIRCUT").Element("NAME")).Trim();
                var dateToCheck = dateEdit_DownloadDate.DateTime;
                string FileName = HaircutFileName.Replace("$date:ddMMyyyy$", dateToCheck.ToString("ddMMyyyy"));/*$"MF_VAR_{dateEdit_DownloadDate.DateTime.STR("ddMMyyyy").UPP()}.csv";*/


                bool filedownloaded = false;

                if (arr_MFundHaircutFolderPath.Length != 0)
                {
                    AddToList("Mutual fund haircut file downloading.");

                    for (int i = 0; i < arr_MFundHaircutFolderPath.Length; i++)
                    {
                        if (!Directory.Exists(arr_MFundHaircutFolderPath[i]))
                            Directory.CreateDirectory(arr_MFundHaircutFolderPath[i]);
                        else
                        {
                            string[] files = Directory.GetFiles(arr_MFundHaircutFolderPath[i], @"*.gz");
                            for (int j = 0; j < files.Count(); j++)
                                File.Delete(files[j]);
                        }
                    }
                    #region API doesn't have MFHaircutfile
                    //try
                    //{
                    //var response = nNSEUtils.Instance.DownloadCommonFile(en_FolderTypes.CM_APPSEC_COLLVAL,FileName, arr_MFundHaircutFolderPath[0]);
                    //    _logger.Debug("DownloadSecurityFile API Response: " + JsonConvert.SerializeObject(response));
                    //    if (response.ResponseStatus == en_ResponseStatus.SUCCESS)
                    //    {
                    //        //DecompressGZAndDelete(new FileInfo(arr_MFundHaircutFolderPath[0] + @"security.gz"), ".txt");
                    //        filedownloaded = true;
                    //    }
                    //}
                    //catch (Exception ee)
                    //{
                    //    _logger.Error(ee);
                    //}
                    #endregion
                    //if (dateToCheck.DayOfWeek == DayOfWeek.Saturday || dateToCheck.DayOfWeek == DayOfWeek.Sunday)
                    //{
                    //    // if so, set the date to the previous Friday
                    //    dateToCheck = (dateToCheck.DayOfWeek == DayOfWeek.Saturday) ? dateToCheck.AddDays(-1) : dateToCheck.AddDays(-2);
                    //}
                    for (int i = 0; i < 7; i++)
                    {
                        FileName = HaircutFileName.Replace("$date:ddMMyyyy$", dateToCheck.ToString("ddMMyyyy"));
                        string URL = arr_MFundHaircutURL[0] + FileName;
                        //trying with website
                        try
                        {
                            if (!filedownloaded)
                            {
                                using (WebClient webClient = new WebClient())
                                {
                                    webClient.DownloadFile(URL, arr_MFundHaircutFolderPath[0] + FileName);
                                    filedownloaded = true;
                                    AddToList($"Mutual Fund Haircut file: {FileName} downloaded successfully.");
                                    MFHaircut = FileName;
                                    break;
                                }
                            }
                        }
                        catch (Exception ee)
                        {
                            _logger.Error(ee,"Downloading MF-HairCut from website");
                            _logger.Debug($"URL passed: {URL} and Savepath selected {arr_MFundHaircutFolderPath[0] + FileName}");
                        }

                        try
                        {
                            if (!filedownloaded)
                            {
                                string filename = arr_MFundHaircutURLlocal[0] + FileName;
                                File.Copy(filename, arr_MFundHaircutFolderPath[0] + FileName, true);
                                MFHaircut = FileName;
                                filedownloaded = true;
                                AddToList($"Mutual Fund Haircut file: {FileName} downloaded successfully.");
                                break;
                            }
                        }
                        catch (Exception ee) 
                        { 
                            _logger.Error(ee,"Copying MF Haircut from Local Folder");
                            _logger.Debug($"Source: {arr_MFundHaircutURLlocal[0] + FileName} and Destination: {arr_MFundHaircutFolderPath[0] + FileName}");
                        }

                        // subtract a day from the date to check the previous day
                        dateToCheck = dateToCheck.AddDays(-1);

                        // skip weekends
                        if (dateToCheck.DayOfWeek == DayOfWeek.Saturday)
                        {
                            dateToCheck = dateToCheck.AddDays(-1);
                        }
                        else if (dateToCheck.DayOfWeek == DayOfWeek.Sunday)
                        {
                            dateToCheck = dateToCheck.AddDays(-2);
                        }
                    }
                    if (!filedownloaded)
                    {
                        AddToList("Mutual Fund Haircut file download failed.", true);
                        return;
                    }

                    if (arr_MFundHaircutFolderPath.Length > 1 && filedownloaded)
                    {
                        // AddToList("NNF Security file downloaded successfully.");
                        for (int i = 1; i < arr_MFundHaircutFolderPath.Length; i++)
                            File.Copy(arr_MFundHaircutFolderPath[0] + FileName, arr_MFundHaircutFolderPath[i] + FileName, true);
                    }
                    if (filedownloaded) { _logger.Debug("Mutual Fund Haircut file downloaded successfully in all Save-Paths."); }
                }
                else
                    AddToList("Invalid path specified for Mutual Fund Haircut file.", true);
            }
            catch (Exception ee)
            {
                _logger.Error(ee); AddToList("Unable to download Mutual Fund Haircut file.", true);
            }
        }

        private void DownloadMCXScrip()
        {
            _logger.Debug($"Checking license status: EnabledSegments.MCX: {_LicenseInfo.EnabledSegments.MCX}"); // Added by Musharraf 10th April 2023
            if (!(_LicenseInfo.EnabledSegments.MCX))
            {
                return;
            }

            try
            {
                _logger.Debug("Executing DownloadMCXScrip():");
                AddToList("Loading MCX ScripFile");
                string MCXFile = ((string)xmlDoc.Element("BOD-Utility").Element("MCX").Element("FILE").Element("NAME")).Trim(); // Filename
                string[] MCXFilePaths = ((string)xmlDoc.Element("BOD-Utility").Element("MCX").Element("FILE").Element("SCRIPFILEPATH")).Split(','); // Array of FilePaths

                string fileNamePattern = MCXFile;
                List<string> latestFiles = new List<string>();

                foreach (string MCXFilePath in MCXFilePaths)
                {
                    string[] files = Directory.GetFiles(MCXFilePath, fileNamePattern);

                    if (files.Length > 0)
                    {
                        string latestFile = files
                            .Select(f => new FileInfo(f))
                            .OrderByDescending(f => f.LastWriteTime)
                            .First()
                            .FullName;

                        latestFiles.Add(latestFile);
                    }
                    else
                    {
                        //AddToList($"Fail to load MCX ScripFile for path: {MCXFilePath}", true);
                        _logger.Debug($"No files found matching the pattern in path: {MCXFilePath}. Check SCRIPFILEPATH or NAME under MCX in Config");
                    }
                }

                if (latestFiles.Count > 0)
                {
                    string latestFile = latestFiles
                        .Select(f => new FileInfo(f))
                        .OrderByDescending(f => f.LastWriteTime)
                        .First()
                        .FullName;

                    _logger.Debug("MCXFile loaded successfully: " + latestFile);
                    AddToList($"MCX File: {latestFile.Substring(latestFile.LastIndexOf('\\') + 1)} loaded successfully");
                    _MCXScripFile = latestFile;
                }
            }
            catch (Exception ee)
            {
                _logger.Error(ee, $"DownloadMCXScrip() {ee}");
            }

            ReadDailMargin();
        }

        private void ReadDailMargin()
        {
            try
            {
                _logger.Debug("Executing  ReadDailMargin():");
                AddToList("Loading DailyMargin");
                string DailyMarginPath = ((string)xmlDoc.Element("BOD-Utility").Element("MCX").Element("FILE").Element("DAILYMARGIN")).Trim(); //Bhavcopy path
                string fileNamePattern = "DailyMargin_*.csv";
                string[] files = Directory.GetFiles(DailyMarginPath, fileNamePattern);

                if (files.Length > 0)
                {
                    string latestFile = files
                        .Select(f => new FileInfo(f))
                        .OrderByDescending(f => f.LastWriteTime)
                        .First()
                        .FullName;

                    _logger.Debug("DailyMargin: " + latestFile);
                    AddToList($"DailyMargin: {latestFile.Substring(latestFile.LastIndexOf('\\') + 1)} loaded successfully");
                }
                else
                {
                    AddToList("Failed to load DailyMargin", true);
                    _logger.Debug($"No files found matching the pattern. Check DailyMargin: {DailyMarginPath} under MCX in Config");
                }
            }
            catch (Exception ee)
            {
                _logger.Error(ee, $"ReadDailMargin() {ee}");
                AddToList("DailyMargin File not found", true);
            }
        }

        private void DownloadMCXBhavcopy()
        {
            _logger.Debug($"Checking license status: EnabledSegments.MCX: {_LicenseInfo.EnabledSegments.MCX}"); // Added by Musharraf 10th April 2023
            if (!(_LicenseInfo.EnabledSegments.MCX))
            {
                return;
            }

            try
            {
                _logger.Debug("Executing DownloadMCXBhavcopy()");
                AddToList("Loading MCX Bhavcopy");
                string MCXBhavCopyPath = ((string)xmlDoc.Element("BOD-Utility").Element("MCX").Element("FILE").Element("BHAVCOPYPATH")).Trim(); //Bhavcopy path
                string fileNamePattern = "BhavCopy_MCXCCL_CO*.csv";


                string[] files = Directory.GetFiles(MCXBhavCopyPath, fileNamePattern);

                if (files.Length > 0)
                {
                    string latestFile = files
                        .Select(f => new FileInfo(f))
                        .OrderByDescending(f => f.LastWriteTime)
                        .First()
                        .FullName;

                    _logger.Debug("MCXBhavcopy: " + latestFile);
                    AddToList($"MCX Bhavcopy: {latestFile.Substring(latestFile.LastIndexOf('\\') + 1)} loaded successfully");
                    _MCXbhavcopy = latestFile.Substring(latestFile.LastIndexOf('\\') + 1);
                }
                else
                {
                    AddToList("Failed to load MCX Bhavcopy", true);
                    _logger.Debug($"No files found matching the pattern. Check BHAVCOPYPATH: {MCXBhavCopyPath} under MCX in Config");
                }
            }
            catch (Exception ee )
            {
                _logger.Error(ee, $"DownloadMCXScrip() {ee}");
                AddToList("MCX Bhavcopy not found", true);
            }
        }

        //added by nikhil | gross F/O
        private void DownloadHaricutFile()
        {
            _logger.Debug($"Checking license status: EnabledSegments.CM and FO: {_LicenseInfo.EnabledSegments.CM} and {_LicenseInfo.EnabledSegments.FO}"); // Added by Musharraf 10th April 2023
            if (!(_LicenseInfo.EnabledSegments.CM && _LicenseInfo.EnabledSegments.FO))
            {
                return;
            }
            try
            {
                _logger.Debug("Executing DownloadHaricutFile()");

                string[] arr_HaircutFolderPath = ds_Config.GET("SAVEPATH", "COLLATERAL-HAIRCUT").SPL(',');
                //string[] arr_HaircutURL = ds_Config.GET("URLs", "COLLATERAL-HAIRCUT").SPL(',');
                string downloadWebsite = ((string)xmlDoc.Element("BOD-Utility").Element("FO").Element("COLLATERAL-HAIRCUT").Element("WEBSITE")).Trim();//To download Website
                string[] arr_HaircutURL = downloadWebsite.Split(',');
                string CollHairCutFile = ((string)xmlDoc.Element("BOD-Utility").Element("FO").Element("COLLATERAL-HAIRCUT").Element("NAME")).Trim();/*$"APPSEC_COLLVAL_{dateEdit_DownloadDate.DateTime.STR("ddMMyyyy").UPP()}.csv";*/

                string Local = ((string)xmlDoc.Element("BOD-Utility").Element("FO").Element("COLLATERAL-HAIRCUT").Element("LOCAL")).Trim();//To download Website
                string[] arr_HaircutURLLocal = Local.Split(',');


                var dateToCheck = dateEdit_DownloadDate.DateTime;


                string FileName = CollHairCutFile.Replace("$date:ddMMyyyy$", dateToCheck.STR("ddMMyyyy"));


                bool fileDownloaded = false;

                if (arr_HaircutFolderPath.Length != 0)
                {
                    AddToList("Collateral Haircut file downloading.");

                    for (int i = 0; i < arr_HaircutFolderPath.Length; i++)
                    {
                        if (!Directory.Exists(arr_HaircutFolderPath[i]))
                            Directory.CreateDirectory(arr_HaircutFolderPath[i]);
                        else
                        {
                            string[] files = Directory.GetFiles(arr_HaircutFolderPath[i], @"APPSEC_COLLVAL_*.csv");
                            for (int j = 0; j < files.Count(); j++)
                                File.Delete(files[j]);
                        }
                    }

                    //if (dateToCheck.DayOfWeek == DayOfWeek.Saturday || dateToCheck.DayOfWeek == DayOfWeek.Sunday)
                    //{
                    //    // if so, set the date to the previous Friday
                    //    dateToCheck = (dateToCheck.DayOfWeek == DayOfWeek.Saturday) ? dateToCheck.AddDays(-1) : dateToCheck.AddDays(-2);
                    //}

                    for (int i = 0; i < 7; i++)
                    {
                        FileName = CollHairCutFile.Replace("$date:ddMMyyyy$", dateToCheck.STR("ddMMyyyy"));

                        try
                        {
                            var response = nNSEUtils.Instance.DownloadCommonFile(en_FolderTypes.CM_APPSEC_COLLVAL, FileName, arr_HaircutFolderPath[0]);
                            _logger.Debug("DownloadAppSecCollFile API Response: " + JsonConvert.SerializeObject(response));
                            if (response.ResponseStatus == en_ResponseStatus.SUCCESS)
                            {
                                //DecompressGZAndDelete(new FileInfo(arr_MFundHaircutFolderPath[0] + @"security.gz"), ".txt");
                                fileDownloaded = true;
                                AddToList($"Collateral Haircut file :{FileName} downloaded successfully.");
                                collateralHaircut = FileName;
                                break;
                            }
                        }
                        catch (Exception ee)
                        {
                            _logger.Error(ee, "Downloading Collateral Haircut File from API:");
                        }


                        string URL = arr_HaircutURL[0] + FileName;
                        //trying with website
                        try
                        {
                            if (!fileDownloaded)
                            {
                                using (WebClient webClient = new WebClient())
                                {
                                    webClient.DownloadFile(URL, arr_HaircutFolderPath[0] + FileName);
                                    fileDownloaded = true;
                                    AddToList($"Collateral Haircut file :{FileName} downloaded successfully.");
                                    collateralHaircut = FileName;
                                    break;
                                }
                            }
                        }
                        catch (Exception ee)
                        {
                            _logger.Error(ee,"Downloading Collateral Haircut File from Website:");
                            _logger.Debug($"Web URL used: {URL} and Savepath: {arr_HaircutFolderPath[0] + FileName}");
                        }
                        #region FTP
                        //if (!fileDownloaded)
                        //{
                        //    //trying with ftp
                        //    try
                        //    {
                        //        using (WebClient webClient = new WebClient())
                        //        {
                        //            //Added to login and download from NSE FTP link. 16MAR2021-Amey
                        //            webClient.Credentials = new NetworkCredential(dict_FTPCred["FO"].Username, dict_FTPCred["FO"].Password);

                        //            webClient.DownloadFile(arr_HaircutURL[0] + FileName, arr_HaircutFolderPath[0] + FileName);

                        //            fileDownloaded = true;
                        //        }
                        //    }
                        //    catch (Exception ee)
                        //    {
                        //        _logger.Error(ee);
                        //    }
                        //}
                        #endregion

                        try
                        {
                            if (!fileDownloaded)
                            {
                                string filename = arr_HaircutURLLocal[0] + FileName;
                                File.Copy(filename, arr_HaircutFolderPath[0] + FileName, true);
                                collateralHaircut = FileName;
                                fileDownloaded = true;
                                AddToList($"Collateral Haircut file :{FileName} downloaded successfully.");
                                break;
                            }
                        }
                        catch (Exception ee) 
                        { 
                            _logger.Error(ee, "Copying Collateral Haircut File from Local Folder:");
                            _logger.Debug($"Source: {arr_HaircutURLLocal[0] + FileName} and Destination: {arr_HaircutFolderPath[0] + FileName}");
                        }

                        // subtract a day from the date to check the previous day
                        dateToCheck = dateToCheck.AddDays(-1);

                        // skip weekends
                        if (dateToCheck.DayOfWeek == DayOfWeek.Saturday)
                        {
                            dateToCheck = dateToCheck.AddDays(-1);
                        }
                        else if (dateToCheck.DayOfWeek == DayOfWeek.Sunday)
                        {
                            dateToCheck = dateToCheck.AddDays(-2);
                        }
                    }
                    if (arr_HaircutFolderPath.Length > 1 && fileDownloaded)
                    {
                        for (int i = 1; i < arr_HaircutFolderPath.Length; i++)
                            File.Copy(arr_HaircutFolderPath[0] + FileName, arr_HaircutFolderPath[i] + FileName, true);
                    }

                    if (fileDownloaded)
                    {
                        _logger.Debug("Collateral Haircut file downloaded successfully in all Save-Paths.");
                    }
                    else
                    {
                        AddToList("Collateral Haircut file download failed.", true);
                    }
                }
                else
                {
                    AddToList("Invalid path specified for Collateral Haircut file.", true);
                }
            }
            catch (Exception ee)
            {
                _logger.Error(ee);
                AddToList("Unable to download Collateral Haircut file.", true);
            }
        }

        //Added by Akshay on 12-10-2021 for Downloading CD Bhavcopy
        private void DownloadCDContractFile()
        {
            _logger.Debug($"Checking license status: EnabledSegments.CD: {_LicenseInfo.EnabledSegments.CD}"); // Added by Musharraf 10th April 2023
            if (!_LicenseInfo.EnabledSegments.CD)
            {
                return;
            }
            try
            {
                _logger.Debug("Executing DownloadCDContractFile():");

                string[] arr_ContractFolderPath = ds_Config.GET("SAVEPATH", "CD_CONTRACT").SPL(',');
                //string[] arr_ContractURL = ds_Config.GET("URLs", "CD_CONTRACT").SPL(',');

                string downloadFTP = ((string)xmlDoc.Element("BOD-Utility").Element("CD").Element("CONTRACT").Element("FTP")).Trim();//To download file FTP
                string[] arr_ContractURL = downloadFTP.Split(',');

                string downloadWebsite = ((string)xmlDoc.Element("BOD-Utility").Element("CD").Element("CONTRACT").Element("WEBSITE")).Trim();//To download Website
                string[] arr_ContractWebURL = downloadWebsite.Split(',');

                string downloadFromLocal = ((string)xmlDoc.Element("BOD-Utility").Element("CD").Element("CONTRACT").Element("LOCAL")).Trim();//To download from local file
                string[] arr_ContractLocalFile = downloadFromLocal.Split(',');

                var FileName = ((string)xmlDoc.Element("BOD-Utility").Element("CD").Element("CONTRACT").Element("NAME")).Trim();//FileName of your Contract file


                bool filedownloaded = false;

                if (arr_ContractFolderPath.Length != 0)
                {
                    AddToList("cd_Contract file downloading.");

                    for (int i = 0; i < arr_ContractFolderPath.Length; i++)
                    {
                        if (!Directory.Exists(arr_ContractFolderPath[i]))
                            Directory.CreateDirectory(arr_ContractFolderPath[i]);
                        else
                        {
                            string[] files = Directory.GetFiles(arr_ContractFolderPath[i], @"NSE_CD_contract_*.csv");
                            for (int j = 0; j < files.Count(); j++)
                                File.Delete(files[j]);
                        }
                    }

                    var dateToCheck = dateEdit_DownloadDate.DateTime;
                    var NSECDContract = FileName.Replace("$date:ddMMyyyy$", dateToCheck.ToString("ddMMyyyy"));
                    // check if today is a weekend day  // Added by Musharraf 3rd April 2023
                    //if (dateToCheck.DayOfWeek == DayOfWeek.Saturday || dateToCheck.DayOfWeek == DayOfWeek.Sunday)
                    //{
                    //    // if so, set the date to the previous Friday
                    //    dateToCheck = (dateToCheck.DayOfWeek == DayOfWeek.Saturday) ? dateToCheck.AddDays(-1) : dateToCheck.AddDays(-2);
                    //}

                    //Added by Musharraf to check previous 7 days files
                    for (int i = 0; i < 7; i++)
                    {
                        // check for the file in the previous 7 working days
                        try
                        {
                            res_General file_check = null;
                            NSECDContract = FileName.Replace("$date:ddMMyyyy$", dateToCheck.ToString("ddMMyyyy"));

                            file_check = nNSEUtils.Instance.DownloadCommonFile(en_FolderTypes.CD_CONTRACT, NSECDContract, arr_ContractFolderPath[0]);
                            _logger.Debug("DownloadCDContract API Response: " + JsonConvert.SerializeObject(file_check));
                            // download the file if it exists
                            if (file_check != null && file_check.ResponseStatus == en_ResponseStatus.SUCCESS)
                            {
                                //_logger.Debug("DownloadCDContractFile API Response: " + JsonConvert.SerializeObject(file_check.Response.ResponseStatus));
                                if (NSECDContract.EndsWith(".gz") && File.Exists(Path.Combine(arr_ContractFolderPath[0], NSECDContract)))
                                {
                                    _logger.Debug("Inside if NSECDContract.EndsWith(.gz) && File.Exists(Path.Combine(arr_ContractFolderPath[0], NSECDContract))");
                                    FileInfo localFile = new FileInfo(arr_ContractFolderPath[0] + NSECDContract);
                                    NSECDContract = DecompressGZAndDelete(localFile, "");
                                    NSECDContract = Path.GetFileName(NSECDContract);
                                }
                                
                                filedownloaded = true;
                                AddToList($"CD Contract file: {NSECDContract} downloaded successfully.");
                                _logger.Debug($"Downloaded CD Contract File From API : {NSECDContract}");
                            }
                            if (filedownloaded)
                            {
                                CD_contract_fileName = NSECDContract;
                                break;
                            }
                        }
                        catch (Exception ee)
                        {
                            _logger.Error(ee,"Downloading CDContract File from API");
                        }
                        #region FTP Download
                        //if (!filedownloaded)
                        //{
                        //    try
                        //    {
                        //        using (WebClient webClient = new WebClient())
                        //        {
                        //            //Added to login and download from NSE FTP link. 16MAR2021-Amey
                        //            webClient.Credentials = new NetworkCredential(dict_FTPCred["CD"].Username, dict_FTPCred["CD"].Password);

                        //            webClient.DownloadFile(arr_ContractURL[0], arr_ContractFolderPath[0] + @"cd_contract.gz");
                        //        }

                        //        File.Delete(arr_ContractFolderPath[0] + @"NSE_CD_contract_" + DateTime.Now.ToString("ddMMyyyy") + ".csv");

                        //        DecompressGZAndDelete(new FileInfo(arr_ContractFolderPath[0] + @"cd_contract.gz"), ".txt");

                        //        filedownloaded = true;
                        //    }
                        //    catch (Exception ee) { _logger.Error(ee, "DownloadExchangeFiles - cd_Contract"); }

                        //}
                        #endregion
                        try
                        {
                                NSECDContract = FileName.Replace("$date:ddMMyyyy$", dateToCheck.ToString("ddMMyyyy"));

                                if (!filedownloaded && !string.IsNullOrEmpty(arr_ContractLocalFile[0]))
                                {
                                    if (NSECDContract.EndsWith(".gz") && File.Exists(Path.Combine(arr_ContractLocalFile[0], NSECDContract)))
                                    {
                                        FileInfo localFile = new FileInfo(arr_ContractLocalFile[0] + NSECDContract);
                                        NSECDContract = DecompressGZAndDelete(localFile, "");
                                        NSECDContract = Path.GetFileName(NSECDContract);
                                    }
                                    else
                                    {
                                        NSECDContract = NSECDContract.Replace(".gz", "");
                                    }

                                    string source = Path.Combine(arr_ContractLocalFile[0], NSECDContract); 
                                    string destination = Path.Combine(arr_ContractFolderPath[0], NSECDContract);
                                    File.Copy(source, destination, true);
                                    filedownloaded = true;
                                    AddToList($"CD Contract file: {NSECDContract} downloaded successfully.");
                                    if (filedownloaded == true)
                                    {
                                        CD_contract_fileName = NSECDContract;
                                        break;
                                    }
                                }
                        }
                        catch (Exception ee)
                        {
                            _logger.Error(ee, "Copying CD Contract file from Localfolder");
                            _logger.Debug($"Source : {Path.Combine(arr_ContractLocalFile[0], NSECDContract)} and Destination: {Path.Combine(arr_ContractFolderPath[0], NSECDContract)}");
                        }
                        // subtract a day from the date to check the previous day
                        dateToCheck = dateToCheck.AddDays(-1);

                        // skip weekends
                        if (dateToCheck.DayOfWeek == DayOfWeek.Saturday)
                        {
                            dateToCheck = dateToCheck.AddDays(-1);
                        }
                        else if (dateToCheck.DayOfWeek == DayOfWeek.Sunday)
                        {
                            dateToCheck = dateToCheck.AddDays(-2);
                        }
                    }
                    //end of the previous 7 days file check
                    if (!filedownloaded)
                    {
                        AddToList("cd_Contract file download failed.", true);
                        return;
                    }

                    if (arr_ContractFolderPath.Length > 1 && filedownloaded)
                    {
                        for (int i = 1; i < arr_ContractFolderPath.Length; i++)
                            File.Copy(arr_ContractFolderPath[0] + CD_contract_fileName, arr_ContractFolderPath[i] + CD_contract_fileName, true);

                        _logger.Debug("cd_Contract file downloaded successfully in all Save-Paths.");
                    }
                }
                else
                    AddToList("Invalid path specified for cd_Contract file.", true);
            }
            catch (Exception ee) { _logger.Error(ee); AddToList("Unable to download cd_Contract file.", true); }
        }

        //added on 30APR2021 by Amey
        private void DownloadBSEScripFile()
        {
            _logger.Debug($"Checking license status: EnabledSegments.CM : {_LicenseInfo.EnabledSegments.CM} ");  // Added by Musharraf 3rd April 2023
            if (!_LicenseInfo.EnabledSegments.CM)
            {
                return;
            }
            try
            {
                _logger.Debug("Executing DownloadBSEScripFile():");

                bool filedownloaded = false;
                string[] arr_SecurityFolderPath = ds_Config.GET("SAVEPATH", "BSECM_SCRIP").SPL(',');
                string downloadFromWeb = ((string)xmlDoc.Element("BOD-Utility").Element("CM").Element("BSE_SCRIP").Element("WEBSITE")).Trim();
                string[] arr_ScripUrl = downloadFromWeb.Split(',');//Links to Download from web
                string downloadFromLocal = ((string)xmlDoc.Element("BOD-Utility").Element("CM").Element("BSE_SCRIP").Element("LOCAL")).Trim();
                string[] arr_Local = downloadFromLocal.Split(',');//Download from Web
                                                                  //string[] arr_ScripUrl = ds_Config.GET("URLs", "BSECM_SCRIP").Split(',');
                string NameFromXml = ((string)xmlDoc.Element("BOD-Utility").Element("CM").Element("BSE_SCRIP").Element("NAME")).Trim(); //Fetch from XML

                string[] arr_OldFiles;
                if (arr_SecurityFolderPath.Length != 0)
                {
                    AddToList("BSE ScripFile file downloading.");

                    var TempDirectory = arr_SecurityFolderPath[0] + "SCRIP\\";
                    if (Directory.Exists(TempDirectory))
                        Directory.Delete(TempDirectory, true);
                    Directory.CreateDirectory(TempDirectory);

                    for (int i = 0; i < arr_SecurityFolderPath.Length; i++)
                    {
                        if (!Directory.Exists(arr_SecurityFolderPath[i]))
                            Directory.CreateDirectory(arr_SecurityFolderPath[i]);
                        else
                        {
                            arr_OldFiles = Directory.GetFiles(arr_SecurityFolderPath[i], @"scrip.zip");
                            for (int j = 0; j < arr_OldFiles.Count(); j++)
                                File.Delete(arr_OldFiles[j]);
                        }
                    }

                    arr_OldFiles = Directory.GetFiles(arr_SecurityFolderPath[0], @"BSE_EQ_SCRIP_*.csv");//changed by musharraf .txt to .csv
                    for (int j = 0; j < arr_OldFiles.Count(); j++)
                        File.Delete(arr_OldFiles[j]);

                    arr_OldFiles = Directory.GetFiles(arr_SecurityFolderPath[0], "index5_*");
                    for (int j = 0; j < arr_OldFiles.Count(); j++)
                        File.Delete(arr_OldFiles[j]);

                    var dateToCheck = dateEdit_DownloadDate.DateTime;

                    try
                    {

                        try
                        {
                            using (WebClient webClient = new WebClient())
                            {
                                webClient.DownloadFile(arr_ScripUrl[0], arr_SecurityFolderPath[0] + @"scrip.zip");
                            }
                            using (ZipFile zip = ZipFile.Read(arr_SecurityFolderPath[0] + @"scrip.zip"))
                                zip.ExtractAll(arr_SecurityFolderPath[0], ExtractExistingFileAction.DoNotOverwrite);

                            File.Delete(arr_SecurityFolderPath[0] + @"scrip.zip");

                            var allFiles = Directory.GetFiles(TempDirectory);
                            foreach (var _File in allFiles)
                            {
                                if (!_File.Contains("BSE_EQ_SCRIP_") && !_File.Contains("index"))
                                    File.Delete(_File);

                            }
                        }
                        catch (Exception ee)
                        {
                            _logger.Error(ee, "Downloading BSEScripFile from Website");
                            _logger.Debug($"BSEScrip File Url: {arr_ScripUrl[0]}");
                        }

                        //Check for latest BSE_EQ_SCRIP from Previous 7 working days
                        //var BSESripFile = @"BSE_EQ_SCRIP_" + dateEdit_DownloadDate.DateTime.STR("ddMMyyyy") + ".csv";

                        var BSESripFile = NameFromXml.Replace("$date:ddMMyyyy$", dateToCheck.ToString("ddMMyyyy"));
                        // check if today is a weekend day
                        //if (dateToCheck.DayOfWeek == DayOfWeek.Saturday || dateToCheck.DayOfWeek == DayOfWeek.Sunday)
                        //{
                        //    // if so, set the date to the previous Friday
                        //    dateToCheck = dateToCheck.AddDays(-(int)dateToCheck.DayOfWeek).AddDays(-1);
                        //}

                        for (int i = 0; i < 7; i++)
                        {
                            BSESripFile = NameFromXml.Replace("$date:ddMMyyyy$", dateToCheck.ToString("ddMMyyyy"));
                           
                            if (File.Exists(Path.Combine(TempDirectory, BSESripFile)))
                            {
                                filedownloaded = true;
                                File.Copy(TempDirectory + BSESripFile/* @"SCRIP.txt"*/, arr_SecurityFolderPath[0] +/*BSESripFile*/BSESripFile, true);
                                AddToList($"BSE ScripFile file: {BSESripFile} downloaded successfully.");

                                try
                                {
                                    var IndexCloseFile = new DirectoryInfo(TempDirectory).GetFiles("index5_*").OrderByDescending(x => x.LastWriteTime).FirstOrDefault();
                                    if (IndexCloseFile != null)
                                    {
                                        File.Copy(IndexCloseFile.FullName, arr_SecurityFolderPath[0] + IndexCloseFile.Name);
                                        AddToList("BSE Index close file downloaded successfully");
                                    }
                                }
                                catch (Exception ee) { _logger.Error(ee); }

                                BSECM_security_fileName = BSESripFile;
                                break;
                            }

                            
                            
                            try
                            {/*Check BOD_Files if the Downloaded file is not present*/
                                if (!filedownloaded)
                                {
                                    File.Copy(arr_Local[0] + BSESripFile/* @"SCRIP.txt"*/, arr_SecurityFolderPath[0] +/*BSESripFile*/BSESripFile, true);
                                    filedownloaded = true;
                                    AddToList($"BSE ScripFile file: {BSESripFile} downloaded successfully.");
                                    BSECM_security_fileName = BSESripFile;
                                    break;
                                }
                            }
                            catch (Exception ee) 
                            { 
                                _logger.Error(ee,"Copying BSEScripFile from Local Folder: ");
                                _logger.Debug($"Source:{arr_Local[0] + BSESripFile} and Destination:{arr_SecurityFolderPath[0] +/*BSESripFile*/BSESripFile}");
                            }
                            // subtract a day from the date to check the previous day
                            dateToCheck = dateToCheck.AddDays(-1);

                            // skip weekends
                            if (dateToCheck.DayOfWeek == DayOfWeek.Saturday)
                            {
                                dateToCheck = dateToCheck.AddDays(-1);
                            }
                            else if (dateToCheck.DayOfWeek == DayOfWeek.Sunday)
                            {
                                dateToCheck = dateToCheck.AddDays(-2);
                            }
                        }
                        #region Commented
                        //Converter here
                        //Convert the file from old to new
                        /*var */
                        //BSESripFile = @"BSE_EQ_SCRIP_" + dateToCheck.STR("ddMMyyyy") + ".csv";
                        //BSESripFile = "BSE_EQ_SCRIP_27032023.csv";

                        //File.Copy(_ScripTXT[0], arr_SecurityFolderPath[0] + SCRIPFILENAME, true);
                        #endregion
                        //filedownloaded = true;
                    }
                    catch (Exception ee)
                    {
                        _logger.Error(ee, "Downloading - BSEScripFile");
                    }

                    #region Copying from BOD_Files
                    //try
                    //{
                    //    if (!filedownloaded)
                    //    {
                    //        File.Copy(arr_Local[0] + $"BSE_EQ_SCRIP_{dateToCheck.STR("ddMMyy")}.txt"/* @"SCRIP.txt"*/, arr_SecurityFolderPath[0] +/*BSESripFile*/$"BSE_EQ_SCRIP_{dateToCheck.STR("ddMMyy")}.txt", true);
                    //        filedownloaded = true;
                    //    }
                    //}
                    //catch (Exception ee) { _logger.Error(ee); }
                    #endregion
                    if (filedownloaded)
                    {
                        if (arr_SecurityFolderPath.Length > 0)
                        {
                            var SCRIPFILENAME = NameFromXml.Replace("$date:ddMMyyyy$", dateToCheck.ToString("ddMMyyyy"));/*$"BSE_EQ_SCRIP_{dateEdit_DownloadDate.DateTime.STR("ddMMyyyy")}.csv";*/

                            for (int i = 1; i < arr_SecurityFolderPath.Length; i++)
                            {
                                arr_OldFiles = Directory.GetFiles(arr_SecurityFolderPath[i], "BSE_EQ_SCRIP_*.csv");
                                for (int j = 0; j < arr_OldFiles.Count(); j++)
                                    File.Delete(arr_OldFiles[j]);

                                File.Copy(arr_SecurityFolderPath[0] + @"SCRIP\" + SCRIPFILENAME/*$"BSE_EQ_SCRIP_{dateEdit_DownloadDate.DateTime.STR("ddMMyyyy")}.csv"*/, arr_SecurityFolderPath[i] + SCRIPFILENAME, true);
                            }
                        }

                        if (Directory.Exists(TempDirectory))
                            Directory.Delete(TempDirectory, true);

                        _logger.Debug("BSE ScripFile file downloaded successfully in all Save-Paths.");
                    }
                    else
                    {
                        AddToList("BSE ScripFile file download failed.", true);

                        if (Directory.Exists(TempDirectory))
                            Directory.Delete(TempDirectory, true);
                    }

                }
                else
                    AddToList("Invalid path specified for BSE ScripFile file.", true);
            }
            catch (Exception ee) { _logger.Error(ee); AddToList("Unable to download BSEScripFile file.", true); }
        }

        private void DownloadFOBhavcopy()
        {
            _logger.Debug($"Checking license status: EnabledSegments.FO: {_LicenseInfo.EnabledSegments.FO}"); // Added by Musharraf 10th April 2023
            if (!_LicenseInfo.EnabledSegments.FO)
            {
                return;
            }

            _logger.Debug("Executing DownloadFOBhavcopy():");
            string[] BhavcopyURL = ((string)xmlDoc.Element("BOD-Utility").Element("FO").Element("BHAVCOPY").Element("WEBSITE")).Trim().Split(','); //website link
            string[] arr_BhavcopyFolderPath = ds_Config.GET("SAVEPATH", "FO_BHAVCOPY").SPL(',');//Saving from web

            try
            {
                string downloadFTP = ((string)xmlDoc.Element("BOD-Utility").Element("FO").Element("BHAVCOPY").Element("FTP")).Trim();//To download file FTP
                string[] arr_FOBhavFTPUrl = downloadFTP.Split(',');

                //BhavCopyURL is downloading it from the website

                string downloadFromLocal = ((string)xmlDoc.Element("BOD-Utility").Element("FO").Element("BHAVCOPY").Element("LOCAL")).Trim();//To download from local file
                string[] arr_BhavFoLocalFile = downloadFromLocal.Split(',');

                var _FileName = ((string)xmlDoc.Element("BOD-Utility").Element("FO").Element("BHAVCOPY").Element("NAME")).Trim();//FileName of your Contract file

                var dateToCheck = dateEdit_DownloadDate.DateTime;
                AddToList($"FO Bhavcopy downloading.");

                bool filedownloaded = false;
                string[] arr_OldBhavcopyFiles;

                for (int i = 1; i < arr_BhavcopyFolderPath.Length; i++)
                {
                    if (!Directory.Exists(arr_BhavcopyFolderPath[i]))
                        Directory.CreateDirectory(arr_BhavcopyFolderPath[i]);
                    else
                    {
                        arr_OldBhavcopyFiles = Directory.GetFiles(arr_BhavcopyFolderPath[i], @"BhavCopy_NSE_FO*.csv");
                        for (int j = 0; j < arr_OldBhavcopyFiles.Count(); j++)
                            File.Delete(arr_OldBhavcopyFiles[j]);
                    }
                }

                var BhavcopyFileName = _FileName.Replace("$date:yyyyMMdd$", dateToCheck.STR("yyyyMMdd"));/*$"fo{dateToCheck.STR("ddMMMyyyy").UPP()}bhav.csv.gz";*/ // Added by Musharraf 3rd April 2023
                var BhavcopyFileNameZip  = BhavcopyFileName+ ".zip";

                arr_OldBhavcopyFiles = Directory.GetFiles(arr_BhavcopyFolderPath[0], @"BhavCopy_NSE_FO*.csv");
                for (int j = 0; j < arr_OldBhavcopyFiles.Count(); j++)
                    File.Delete(arr_OldBhavcopyFiles[j]);

                var FOBhavcopyFilename = string.Empty;
                
                // Added by Musharraf 3rd April 2023
                //if (dateToCheck.DayOfWeek == DayOfWeek.Saturday || dateToCheck.DayOfWeek == DayOfWeek.Sunday)
                //{
                //    // if so, set the date to the previous Friday
                //    dateToCheck = (dateToCheck.DayOfWeek == DayOfWeek.Saturday) ? dateToCheck.AddDays(-1) : dateToCheck.AddDays(-2);
                //}

                //Start for previous 7 days logic
                for (int i = 0; i < 7; i++)
                {
                    //BhavcopyFileName = _FileName.Replace("$date:ddMMyyyy$", dateToCheck.STR("ddMMyyyy"));//used in all 3 methods to download

                    try
                    {
                        var response = nNSEUtils.Instance.DownloadCommonFile(en_FolderTypes.FO_BHAVCOPY, BhavcopyFileNameZip, arr_BhavcopyFolderPath[0]);
                        _logger.Debug("DownloadFOBhavcopyFile API Response: " + JsonConvert.SerializeObject(response));
                        if (response.ResponseStatus == en_ResponseStatus.SUCCESS)
                        {
                            using (ZipFile zip = ZipFile.Read(arr_BhavcopyFolderPath[0] + BhavcopyFileNameZip))
                                zip.ExtractAll(arr_BhavcopyFolderPath[0], ExtractExistingFileAction.DoNotOverwrite);

                            File.Delete(arr_BhavcopyFolderPath[0] + $"{BhavcopyFileName}.zip");

                            //FOBhavcopyFilename = DecompressGZAndDelete(new FileInfo(arr_BhavcopyFolderPath[0] + BhavcopyFileName), string.Empty/*".csv"*/);
                            filedownloaded = true;
                            AddToList($"FO Bhavcopy: {BhavcopyFileName.Substring(0, BhavcopyFileName.LastIndexOf(".csv") + 4)} downloaded successfully.");
                            _FOBhavcopy = BhavcopyFileName;
                            break;
                        }
                    }
                    catch (Exception ee)
                    {
                        _logger.Error(ee,"Downloading FO Bhavcopy from API:");
                    }

                    //Downloading from website
                    if (!filedownloaded)
                    {
                        try
                        {
                            string url =  BhavcopyURL[0] + BhavcopyFileNameZip;
                            //$"NSE_FO_bhavcopy_{dateToCheck.ToString("ddMMyyyy")}.csv";/*$"{dateEdit_DownloadDate.DateTime.STR("yyyy")}/{dateEdit_DownloadDate.DateTime.STR("MMM").UPP()}/fo{dateToCheck.STR("ddMMMyyyy").UPP()}bhav.csv.zip";*/
                            //BhavcopyFileName = $"NSE_FO_bhavcopy_{dateToCheck.ToString("ddMMyyyy")}.csv";
                            using (WebClient webClient = new WebClient())
                            {
                                webClient.DownloadFile(url, arr_BhavcopyFolderPath[0] + BhavcopyFileNameZip);
                            }
                            using (ZipFile zip = ZipFile.Read(arr_BhavcopyFolderPath[0] + BhavcopyFileNameZip))
                                zip.ExtractAll(arr_BhavcopyFolderPath[0], ExtractExistingFileAction.DoNotOverwrite);
                            
                            File.Delete(arr_BhavcopyFolderPath[0] + $"{BhavcopyFileName}.zip");
                            //File.Delete(arr_BhavcopyFolderPath[0] + BhavcopyFileName.Replace(".gz", ".zip"));
                            filedownloaded = true;
                            AddToList($"FO Bhavcopy: {BhavcopyFileName} downloaded successfully.");
                            _FOBhavcopy = BhavcopyFileName; //BhavcopyFileName.Substring(0, BhavcopyFileName.LastIndexOf(".csv") + 4);
                            break;
                        }
                        catch (Exception ee) 
                        { 
                            _logger.Error(ee,"Downloading FO Bhavcopy from Website:"); 
                            _logger.Debug($"Url Passed: {BhavcopyURL[0] + BhavcopyFileName}");
                        }
                    }

                    #region FTP
                    //if (!filedownloaded)
                    //    {

                    //        //Downloading from FTP
                    //        try
                    //        {
                    //            using (WebClient webClient = new WebClient())
                    //            {
                    //                //Added to login and download from NSE FTP link. 16MAR2021-Amey
                    //                webClient.Credentials = new NetworkCredential(dict_FTPCred["FO"].Username, dict_FTPCred["FO"].Password);

                    //                webClient.DownloadFile(arr_FOBhavFTPUrl[0] + BhavcopyFileName, arr_BhavcopyFolderPath[0] + BhavcopyFileName);

                    //            }

                    //            FOBhavcopyFilename = DecompressGZAndDelete(new FileInfo(arr_BhavcopyFolderPath[0] + BhavcopyFileName), ".csv");
                    //            filedownloaded = true;
                    //            AddToList($"FO Bhavcopy: {FOBhavcopyFilename} downloaded successfully.");
                    //            break;

                    //        }
                    //        catch (Exception ee) { _logger.Error(ee); }
                    //    }
                    #endregion

                    if (!filedownloaded)
                    {
                        try
                        {
                            //var FileName = (_FileName.Replace("$date:ddMMMyyyy$", dateToCheck.STR("ddMMMyyyy").UPP())).Substring(0, _FileName.Replace("$date:ddMMMyyyy$", dateToCheck.STR("ddMMMyyyy").UPP()).LastIndexOf(".csv") + 4);/*$"fo{dateToCheck.STR("ddMMMyyyy").UPP()}bhav.csv";*/
                            File.Copy(arr_BhavFoLocalFile[0] + BhavcopyFileName, arr_BhavcopyFolderPath[0] + BhavcopyFileName, true);

                            AddToList($"FO Bhavcopy: {BhavcopyFileName.Substring(0, BhavcopyFileName.LastIndexOf(".csv") + 4)} downloaded successfully.");
                            _FOBhavcopy = BhavcopyFileName;
                            filedownloaded = true;
                            break;
                        }
                        catch (Exception ee) 
                        { 
                            _logger.Error(ee,"Copying FO Bhavcopy from Local Folder:");
                            _logger.Debug($"Source : {arr_BhavFoLocalFile[0] + BhavcopyFileName} and Destination: {arr_BhavcopyFolderPath[0] + BhavcopyFileName}");
                        }
                    }
                    // subtract a day from the date to check the previous day
                    dateToCheck = dateToCheck.AddDays(-1);

                    // skip weekends
                    if (dateToCheck.DayOfWeek == DayOfWeek.Saturday)
                    {
                        dateToCheck = dateToCheck.AddDays(-1);
                    }
                    else if (dateToCheck.DayOfWeek == DayOfWeek.Sunday)
                    {
                        dateToCheck = dateToCheck.AddDays(-2);
                    }
                }

                //end of the the logic for previous 7 days 

                if (!filedownloaded)
                {
                    AddToList("FO Bhavcopy file download failed.", true);

                    return;
                }
                else
                {
                    if (arr_BhavcopyFolderPath.Length > 1)
                    {
                        var arr_BhavcopyCSV = Directory.GetFiles(arr_BhavcopyFolderPath[0], "BhavCopy_NSE_FO_*.csv");
                        //BhavcopyFileName = _FileName.Replace("$date:ddMMMyyyy$", dateToCheck.STR("ddMMMyyyy").UPP());/*$"fo{dateEdit_DownloadDate.DateTime.STR("ddMMMyyyy").UPP()}bhav.csv";*/

                        for (int i = 1; i < arr_BhavcopyFolderPath.Length; i++)
                        {
                            ////var oldcsvFiles = Directory.GetFiles(arr_BhavcopyFolderPath[i], "fo*.csv");
                            //var oldcsvFiles = new DirectoryInfo(arr_BhavcopyFolderPath[i]).GetFiles("BhavCopy_NSE_FO_*.csv").OrderByDescending(file => file.LastWriteTime).Select(file => file.FullName).ToArray();

                            //for (int j = 0; j < oldcsvFiles.Count(); j++)
                            //    File.Delete(oldcsvFiles[j]);

                            if (arr_BhavcopyCSV.Length > 0 && File.Exists(arr_BhavcopyCSV[0]))
                            {
                                File.Copy(arr_BhavcopyCSV[0], arr_BhavcopyFolderPath[i] + BhavcopyFileName, true);
                            }
                        }
                    }

                    _logger.Debug($"FO Bhavcopy downloaded successfully in all Save-Paths.");
                }


            }
            catch (Exception ee)
            {
                _logger.Error(ee, $"Download FO Bhavcopy With {BhavcopyURL}");
                AddToList($"Unable to download FO Bhavcopy file.", true);
            }
        }

        private void DownloadCMBhavcopy(string[] BhavcopyURL, string[] arr_BhavcopyFolderPath)
        {
            _logger.Debug($"Checking license status: EnabledSegments.CM: {_LicenseInfo.EnabledSegments.CM}"); // Added by Musharraf 3rd April 2023
            if (!_LicenseInfo.EnabledSegments.CM)
            {
                return;
            }
            try
            {
                _logger.Debug("Executing DownloadCMBhavcopy()");
                AddToList($"CM Bhavcopy downloading.");

                bool filedownloaded = false;

                string[] arr_OldBhavcopyFiles;

                for (int i = 0; i < arr_BhavcopyFolderPath.Length; i++)
                {
                    if (!Directory.Exists(arr_BhavcopyFolderPath[i]))
                        Directory.CreateDirectory(arr_BhavcopyFolderPath[i]);
                    else
                    {
                        arr_OldBhavcopyFiles = Directory.GetFiles(arr_BhavcopyFolderPath[i], @"BhavCopy_NSE_CM_*.csv");
                        for (int j = 0; j < arr_OldBhavcopyFiles.Count(); j++)
                            File.Delete(arr_OldBhavcopyFiles[j]);
                    }
                }

                string downloadFTP = ((string)xmlDoc.Element("BOD-Utility").Element("CM").Element("BHAVCOPY").Element("FTP")).Trim();//To download file FTP
                string[] arr_ContractURL = downloadFTP.Split(',');

                string downloadWebsite = ((string)xmlDoc.Element("BOD-Utility").Element("CM").Element("BHAVCOPY").Element("WEBSITE")).Trim();//To download Website
                BhavcopyURL = downloadWebsite.Split(',');

                string downloadFromLocal = ((string)xmlDoc.Element("BOD-Utility").Element("CM").Element("BHAVCOPY").Element("LOCAL")).Trim();//To download from local file
                string[] arr_ContractLocalFile = downloadFromLocal.Split(',');

                var FileName = ((string)xmlDoc.Element("BOD-Utility").Element("CM").Element("BHAVCOPY").Element("NAME")).Trim();//FileName of your Contract file

                //var BhavcopyFileName = "NSE_CM_bhavcopy_" + dateEdit_DownloadDate.DateTime.STR("ddMMyyyy") + ".csv"; //$"{dateEdit_DownloadDate.DateTime.STR("ddMM")}0000.md";
                #region Previous 7 day logic                                                                                                     
                //Added by Musharraf to check previous 7 days files // Added by Musharraf 3rd April 2023
                
                var dateToCheck = dateEdit_DownloadDate.DateTime;
                var BhavcopyFileName = FileName.Replace("$date:yyyyMMdd$", dateToCheck.ToString("yyyyMMdd"));
                var BhavcopyFileNameZIP = BhavcopyFileName + ".zip";
                // check if today is a weekend day
                //if (dateToCheck.DayOfWeek == DayOfWeek.Saturday || dateToCheck.DayOfWeek == DayOfWeek.Sunday)
                //{
                //    // if so, set the date to the previous Friday
                //    dateToCheck = dateToCheck.AddDays(-(int)dateToCheck.DayOfWeek).AddDays(-1);
                //}

                // check for the file in the previous 7 working days
                for (int j = 0; j < 7; j++)
                {
                    //BhavcopyFileName = FileName.Replace("$date:yyyyMMdd$", dateToCheck.ToString("yyyyMMdd"));
                    //BhavcopyFileName = "NSE_CM_bhavcopy_" + dateToCheck.ToString("ddMMyyyy") + ".csv";
                    // download the file if it exists
                    try
                    {
                        var APIFileName = Path.ChangeExtension(BhavcopyFileName, " ").Substring(0, Path.ChangeExtension(BhavcopyFileName, " ").LastIndexOf('.'));

                        var response = nNSEUtils.Instance.DownloadCommonFile(en_FolderTypes.CM_BHAVCOPY, BhavcopyFileNameZIP, arr_BhavcopyFolderPath[0]);

                        _logger.Debug("DownloadCMBhavcopyFile API Response: " + JsonConvert.SerializeObject(response));
                        if (response.ResponseStatus == en_ResponseStatus.SUCCESS)
                        {
                            using (ZipFile zip = ZipFile.Read(arr_BhavcopyFolderPath[0] + BhavcopyFileNameZIP))
                            {
                                zip.ExtractAll(arr_BhavcopyFolderPath[0], ExtractExistingFileAction.DoNotOverwrite);
                            }

                            File.Delete(arr_BhavcopyFolderPath[0] + $"{BhavcopyFileName}.zip");

                            AddToList($"CM Bhavcopy: {APIFileName} downloaded successfully.");

                            _NSE_CM_bhavcopy = BhavcopyFileName;
                            filedownloaded = true;
                            
                        }
                    }
                    catch (Exception ee)
                    {
                        _logger.Error(ee,"Downloading CM Bhavcopy Using API");
                    }

                    try
                    {
                        if (!filedownloaded)
                        {
                            //BhavcopyFileName = $"NSE_CM_bhavcopy_{dateToCheck.STR("ddMMyyyy")}.csv";//Changed to datetocheck
                            //BhavcopyFileName = FileName.Replace("$date:ddMMyyyy$", dateToCheck.ToString("ddMMyyyy"));
                            string url = BhavcopyURL[0] + BhavcopyFileNameZIP ; //$"{dateToCheck.STR("yyyy")}/{dateToCheck.STR("MMM").UPP()}/{BhavcopyFileName}";/*NSE_CM_bhavcopy_{dateToCheck.STR("ddMMyyyy")}.csv.zip*/

                            using (WebClient webClient = new WebClient())
                            {
                                webClient.DownloadFile(url, arr_BhavcopyFolderPath[0] + BhavcopyFileNameZIP/*$"NSE_CM_bhavcopy_{dateToCheck.STR("ddMMyyyy")}.csv.zip"*/);
                            }

                            /*$"NSE_CM_bhavcopy_{dateToCheck.STR("ddMMyyyy")}.csv.zip"*/
                            using (ZipFile zip = ZipFile.Read(arr_BhavcopyFolderPath[0] + BhavcopyFileNameZIP))
                            {
                                zip.ExtractAll(arr_BhavcopyFolderPath[0], ExtractExistingFileAction.DoNotOverwrite);
                            }
                            
                            File.Delete(arr_BhavcopyFolderPath[0] + $"{BhavcopyFileName}.zip");
                            filedownloaded = true;
                            if (filedownloaded == true)
                            {
                                //BhavcopyFileName = Path.ChangeExtension(BhavcopyFileName, " ").Substring(0, Path.ChangeExtension(BhavcopyFileName, " ").LastIndexOf('.'));
                                _NSE_CM_bhavcopy = BhavcopyFileName;
                                AddToList($"CM Bhavcopy: {BhavcopyFileName} downloaded successfully.");
                                break;
                            }
                        }
                    }
                    catch (Exception ee) 
                    { 
                        _logger.Error(ee,"Downloading CM Bhavcopy Using Website");
                        _logger.Debug($"Web URL passed: {BhavcopyURL[0]} + {dateToCheck.STR("yyyy")}/{dateToCheck.STR("MMM").UPP()}/{BhavcopyFileName}");
                    }

                    #region FTP
                    //if (!filedownloaded)
                    //    {
                    //        try
                    //        {
                    //            using (WebClient webClient = new WebClient())
                    //            {
                    //                //Added to login and download from NSE FTP link. 16MAR2021-Amey
                    //                webClient.Credentials = new NetworkCredential(dict_FTPCred["GUEST"].Username, dict_FTPCred["GUEST"].Password);

                    //                webClient.DownloadFile(BhavcopyURL[0] + BhavcopyFileName, arr_BhavcopyFolderPath[0] + BhavcopyFileName);
                    //            }


                    //            filedownloaded = true;
                    //            AddToList($"CM Bhavcopy: {BhavcopyFileName.Substring(0, BhavcopyFileName.LastIndexOf(".csv") + 4)} downloaded successfully.");
                    //            if (filedownloaded == true)
                    //            {
                    //                _NSE_CM_bhavcopy = BhavcopyFileName;
                    //                break;
                    //            }
                    //        }
                    //        catch (Exception ee) { _logger.Error(ee); }

                    //    }
                    #endregion
                    try
                    {
                        if (!filedownloaded)
                        {
                            //BhavcopyFileName = Path.ChangeExtension(BhavcopyFileName, " ").Substring(0, Path.ChangeExtension(BhavcopyFileName, " ").LastIndexOf('.'));
                            string filename = arr_ContractLocalFile[0] + BhavcopyFileName;
                            File.Copy(filename, arr_BhavcopyFolderPath[0] + BhavcopyFileName, true);
                            _NSE_CM_bhavcopy = BhavcopyFileName;
                            filedownloaded = true;
                            AddToList($"CM Bhavcopy: {BhavcopyFileName} downloaded successfully.");
                            break;
                        }
                    }
                    catch (Exception ee) { _logger.Error(ee,"Copying CM Bhavcopy from Local Folder:"); }


                    // subtract a day from the date to check the previous day
                    dateToCheck = dateToCheck.AddDays(-1);

                    // skip weekends
                    if (dateToCheck.DayOfWeek == DayOfWeek.Saturday)
                    {
                        dateToCheck = dateToCheck.AddDays(-1);
                    }
                    else if (dateToCheck.DayOfWeek == DayOfWeek.Sunday)
                    {
                        dateToCheck = dateToCheck.AddDays(-2);
                    }
                }

                //end of the previous 7 days file check
                #endregion

                if (!filedownloaded)
                {
                    AddToList("CM Bhavcopy file download failed.", true);

                    return;
                }
                else
                {
                    if (arr_BhavcopyFolderPath.Length > 1)
                    {
                        var arr_BhavcopyCSV = Directory.GetFiles(arr_BhavcopyFolderPath[0], "BhavCopy_NSE_CM_*.csv");

                        for (int i = 1; i < arr_BhavcopyFolderPath.Length; i++)
                        {
                            //var oldcsvFiles = Directory.GetFiles(arr_BhavcopyFolderPath[i], "BhavCopy_NSE_CM_*.csv");
                            //for (int j = 1; j < oldcsvFiles.Count(); j++)
                            //    File.Delete(oldcsvFiles[j]);

                            if (arr_BhavcopyCSV.Length > 0 && File.Exists(arr_BhavcopyCSV[0]))
                            {
                                File.Copy(arr_BhavcopyCSV[0], arr_BhavcopyFolderPath[i] + BhavcopyFileName, true);
                            }
                        }
                    }
                }

                _logger.Debug($"CM Bhavcopy downloaded successfully in all Save-Paths.");
            }
            catch (Exception ee)
            {
                _logger.Error(ee, $"Download CM Bhavcopy With {BhavcopyURL}");
                AddToList($"Unable to download CM Bhavcopy file.", true);
            }
        }

        //Added by Akshay on 12-10-2021 for Downloading CD Bhavcopy
        private void DownloadCDBhavcopy(string[] BhavcopyURL, string[] arr_BhavcopyFolderPath)
        {
            _logger.Debug($"Checking license status: EnabledSegments.CD: {_LicenseInfo.EnabledSegments.CD}");// Added by Musharraf 10th April 2023
            if (!_LicenseInfo.EnabledSegments.CD)
            {
                return;
            }
            try
            {
                _logger.Debug("Executing DownloadCDBhavcopy(): ");
                AddToList($"CD Bhavcopy downloading.");

                bool filedownloaded = false;

                string[] arr_OldBhavcopyFiles;

                for (int i = 1; i < arr_BhavcopyFolderPath.Length; i++)
                {
                    if (!Directory.Exists(arr_BhavcopyFolderPath[i]))
                        Directory.CreateDirectory(arr_BhavcopyFolderPath[i]);
                    else
                    {
                        arr_OldBhavcopyFiles = Directory.GetFiles(arr_BhavcopyFolderPath[i], @"BhavCopy_NSE_CD_*.csv");
                        for (int j = 0; j < arr_OldBhavcopyFiles.Count(); j++)
                            File.Delete(arr_OldBhavcopyFiles[j]);
                    }
                }

                string downloadFTP = ((string)xmlDoc.Element("BOD-Utility").Element("CD").Element("BHAVCOPY").Element("FTP")).Trim();//To download file FTP
                string[] arr_SecurityUrl = downloadFTP.Split(',');

                //BhavcopyURL is were the website is

                string downloadFromLocal = ((string)xmlDoc.Element("BOD-Utility").Element("CD").Element("BHAVCOPY").Element("LOCAL")).Trim();//To download from local file
                string[] arr_cdBhavLocalFile = downloadFromLocal.Split(',');

                var FileName = ((string)xmlDoc.Element("BOD-Utility").Element("CD").Element("BHAVCOPY").Element("NAME")).Trim();//FileName of your file
                var dateToCheck = dateEdit_DownloadDate.DateTime;
                var BhavcopyFileName = FileName.Replace("$date:ddMMyyyy$", dateToCheck.STR("ddMMyyyy")); /*$"FINAL_{dateEdit_DownloadDate.DateTime.STR("ddMM")}0000.md";*/

                //arr_OldBhavcopyFiles = Directory.GetFiles(arr_BhavcopyFolderPath[0], @"FINAL*.csv");
                //for (int j = 0; j < arr_OldBhavcopyFiles.Count(); j++)
                //    File.Delete(arr_OldBhavcopyFiles[j]);

                // Added by Musharraf 5th April 2023
                //if (dateToCheck.DayOfWeek == DayOfWeek.Saturday || dateToCheck.DayOfWeek == DayOfWeek.Sunday)
                //{
                //    // if so, set the date to the previous Friday
                //    dateToCheck = (dateToCheck.DayOfWeek == DayOfWeek.Saturday) ? dateToCheck.AddDays(-1) : dateToCheck.AddDays(-2);
                //}
                for (int i = 0; i < 7; i++)
                {
                    BhavcopyFileName = FileName.Replace("$date:yyyyMMdd$", dateToCheck.STR("yyyyMMdd"));
                    var BhavcopyFileNameZip = BhavcopyFileName + ".zip";// FileName.Replace("$date:yyyyMMdd$", dateToCheck.STR("yyyyMMdd"));
                    try
                    {
                        var response = nNSEUtils.Instance.DownloadCommonFile(en_FolderTypes.CD_BHAVCOPY, BhavcopyFileNameZip, arr_BhavcopyFolderPath[0]);
                        _logger.Debug("DownloadCDBhavcopyFile API Response: " + JsonConvert.SerializeObject(response));
                        if (response.ResponseStatus == en_ResponseStatus.SUCCESS)
                        {
                          
                            filedownloaded = true;
                            AddToList($"CD Bhavcopy :{BhavcopyFileName} downloaded successfully.");
                            _CDBhavcopy = BhavcopyFileName;
                            break;
                        }
                    }
                    catch (Exception ee)
                    {
                        _logger.Error(ee,"Downloading CD Bhavcopy using API:");
                    }


                    try
                    {
                        if (!filedownloaded)
                        {
                            string url =  BhavcopyURL[0] + BhavcopyFileNameZip;/*$"{dateEdit_DownloadDate.DateTime.STR("yyyy")}/{dateEdit_DownloadDate.DateTime.STR("MMM").UPP()}/cd{dateEdit_DownloadDate.DateTime.STR("ddMMMyyyy").UPP()}bhav.csv.zip";*/

                            using (WebClient webClient = new WebClient())
                            {
                                webClient.DownloadFile(url, arr_BhavcopyFolderPath[0] + BhavcopyFileNameZip/*$"cd{dateEdit_DownloadDate.DateTime.STR("ddMMMyyyy").UPP()}bhav.zip"*/);
                               
                                using (ZipFile zip = ZipFile.Read(arr_BhavcopyFolderPath[0] + BhavcopyFileNameZip))
                                    zip.ExtractAll(arr_BhavcopyFolderPath[0], ExtractExistingFileAction.DoNotOverwrite);

                                File.Delete(arr_BhavcopyFolderPath[0] + BhavcopyFileNameZip);

                                filedownloaded = true;

                                AddToList($"CD Bhavcopy downloaded successfully.");
                                _CDBhavcopy = BhavcopyFileName;
                               
                            }

                          
                        }
                    }
                    catch (Exception ee) { _logger.Error(ee, "Downloading CD Bhavcopy using Website:"); }
                    #region FTP
                    //if (!filedownloaded)
                    //{
                    //    try
                    //    {
                    //        using (WebClient webClient = new WebClient())
                    //        {
                    //            //Added to login and download from NSE FTP link. 16MAR2021-Amey
                    //            webClient.Credentials = new NetworkCredential(dict_FTPCred["CD"].Username, dict_FTPCred["CD"].Password);

                    //            webClient.DownloadFile(BhavcopyURL[0] + BhavcopyFileName, arr_BhavcopyFolderPath[0] + BhavcopyFileName);

                    //        }


                    //        filedownloaded = true;
                    //    }
                    //    catch (Exception ee) { _logger.Error(ee); }

                    //}
                    #endregion
                    try
                    {
                        if (!filedownloaded)
                        {
                            string sourcePath = Path.Combine(arr_cdBhavLocalFile[0], BhavcopyFileName);
                            string destinationPath = Path.Combine(arr_BhavcopyFolderPath[0], BhavcopyFileName);
                            File.Copy(sourcePath, destinationPath, true);

                            filedownloaded = true;
                            AddToList($"CD Bhavcopy :{BhavcopyFileName.Substring(0, BhavcopyFileName.LastIndexOf(".csv") + 4)} downloaded successfully.");
                            _CDBhavcopy = BhavcopyFileName;
                            break;
                        }
                    }
                    catch (Exception ee) { _logger.Error(ee,"Copying CD Bhavcopy from Local Folder: "); }

                    // subtract a day from the date to check the previous day
                    dateToCheck = dateToCheck.AddDays(-1);

                    // skip weekends
                    if (dateToCheck.DayOfWeek == DayOfWeek.Saturday)
                    {
                        dateToCheck = dateToCheck.AddDays(-1);
                    }
                    else if (dateToCheck.DayOfWeek == DayOfWeek.Sunday)
                    {
                        dateToCheck = dateToCheck.AddDays(-2);
                    }
                }
                //}


                if (!filedownloaded)
                {
                    AddToList("CD Bhavcopy file download failed.", true);

                    return;
                }
                else
                {
                    if (arr_BhavcopyFolderPath.Length > 1)
                    {
                        var arr_BhavcopyCSV = Directory.GetFiles(arr_BhavcopyFolderPath[0], "BhavCopy_NSE_CD_*.csv");
                        for (int i = 1; i < arr_BhavcopyFolderPath.Length; i++)
                        {
                            //var oldcsvFiles = new DirectoryInfo(arr_BhavcopyFolderPath[i]).GetFiles("BhavCopy_NSE_CD_*.csv").OrderByDescending(file => file.LastWriteTime).Select(file => file.FullName).ToArray();
                            //for (int j = 1; j < oldcsvFiles.Count(); j++)
                            //    File.Delete(oldcsvFiles[j]);

                            if (arr_BhavcopyCSV.Length > 0 && File.Exists(arr_BhavcopyCSV[0]))
                            {
                                File.Copy(arr_BhavcopyCSV[0], arr_BhavcopyFolderPath[i] + BhavcopyFileName, true);
                            }
                        }
                    }
                }

                _logger.Debug($"CD Bhavcopy downloaded successfully in all Save-Paths.");
            }
            catch (Exception ee) { _logger.Error(ee, $"Download CD Bhavcopy With {BhavcopyURL}"); AddToList($"Unable to download CD Bhavcopy file.", true); }
        }

        // Added by SNehadri on 08NOV2021
        private void DownloadSnapShot(string[] arr_DailySnapShotUrl, string[] arr_DailySnapShotFilePath)
        {
            _logger.Debug($"Checking license status: EnabledSegments.FO: {_LicenseInfo.EnabledSegments.FO}");  // Added by Musharraf 10th April 2023
            if (!_LicenseInfo.EnabledSegments.FO)
            {
                return;
            }
            try
            {
                _logger.Debug("Executing DownloadSnapShot()");
                AddToList("Daily Snapshot Downloading");

                //Website URl is passed in function parametere
                string downloadFromLocal = ((string)xmlDoc.Element("BOD-Utility").Element("FO").Element("DAILY_SNAPSHOT").Element("LOCAL")).Trim();//To download from local file
                string[] arr_SnapLocalFile = downloadFromLocal.Split(',');
                var dateToCheck = dateEdit_DownloadDate.DateTime;
                string FileName = ((string)xmlDoc.Element("BOD-Utility").Element("FO").Element("DAILY_SNAPSHOT").Element("NAME")).Trim();
                string SnapShotFileName = FileName.Replace("$date:ddMMyyyy$", dateToCheck.STR("ddMMyyyy"));/*$"ind_close_all_{dateEdit_DownloadDate.DateTime.STR("ddMMyyyy")}.csv";*/
                bool filedownloaded = false;

                if (!File.Exists(arr_DailySnapShotFilePath[0] + SnapShotFileName))
                {   // Added by Musharraf 6th April 2023
                    //if (dateToCheck.DayOfWeek == DayOfWeek.Saturday || dateToCheck.DayOfWeek == DayOfWeek.Sunday)
                    //{
                    //    // if so, set the date to the previous Friday
                    //    dateToCheck = (dateToCheck.DayOfWeek == DayOfWeek.Saturday) ? dateToCheck.AddDays(-1) : dateToCheck.AddDays(-2);
                    //}
                    for (int i = 0; i < 7; i++)
                    {
                        SnapShotFileName = FileName.Replace("$date:ddMMyyyy$", dateToCheck.STR("ddMMyyyy"));
                        try
                        {
                            using (WebClient webClient = new WebClient())
                            {
                                webClient.DownloadFile(arr_DailySnapShotUrl[0] + SnapShotFileName, arr_DailySnapShotFilePath[0] + SnapShotFileName);
                            }
                            filedownloaded = true;
                            AddToList($"Daily SnapShot file:{SnapShotFileName} downloaded successfully");
                            DailySnapshot = SnapShotFileName;
                            break;
                        }
                        catch (Exception ee) { _logger.Error(ee,"Downloading Daily Snap Shot from Web"); }

                        try
                        {
                            if (!filedownloaded)
                            {
                                File.Copy(arr_SnapLocalFile[0] + SnapShotFileName, arr_DailySnapShotFilePath[0] + SnapShotFileName, true);
                                filedownloaded = true;
                                AddToList($"Daily SnapShot file:{SnapShotFileName} downloaded successfully");
                                DailySnapshot = SnapShotFileName;
                                break;
                            }
                        }
                        catch (Exception ee) { _logger.Error(ee, "Copying Daily Snapshot From local Folder"); }
                        // subtract a day from the date to check the previous day
                        dateToCheck = dateToCheck.AddDays(-1);

                        // skip weekends
                        if (dateToCheck.DayOfWeek == DayOfWeek.Saturday)
                        {
                            dateToCheck = dateToCheck.AddDays(-1);
                        }
                        else if (dateToCheck.DayOfWeek == DayOfWeek.Sunday)
                        {
                            dateToCheck = dateToCheck.AddDays(-2);
                        }
                    }

                    if (filedownloaded)
                    {
                        var arr_Lines = File.ReadAllLines(arr_DailySnapShotFilePath[0] + SnapShotFileName);

                        for (int i = 0; i < arr_DailySnapShotFilePath.Length; i++)
                        {
                            var oldcsvFiles = Directory.GetFiles(arr_DailySnapShotFilePath[i], "ind_close_all_*.csv");
                            for (int j = 1; j < oldcsvFiles.Count(); j++)
                                File.Delete(oldcsvFiles[j]);

                            File.WriteAllLines(arr_DailySnapShotFilePath[i] + SnapShotFileName, arr_Lines);
                        }
                        arr_Lines = null;

                        _logger.Debug("Daily SnapShot file downloaded successfully in all Save-Paths");
                    }
                    else
                        AddToList("Daily SnapShot failed to download", true);
                }
                else
                {
                    AddToList($"Daily SnapShot {SnapShotFileName} successfully downloaded");
                    DailySnapshot = SnapShotFileName;
                }
            }
            catch (Exception ee) { _logger.Error(ee, "Download DailySnap Shot"); AddToList("Daily SnapShot failed to download", true); }
        }

        //added on 30APR2021 by Amey
        private void DownloadBSECMBhavcopy(string[] BhavcopyURL, string[] arr_BhavcopyFolderPath)
        {
            _logger.Debug($"Checking license status: EnabledSegments.CM: {_LicenseInfo.EnabledSegments.CM}"); // Added by Musharraf 10th April 2023
            if (!_LicenseInfo.EnabledSegments.CM)
            {
                return;
            }
            try
            {
                _logger.Debug("Executing DownloadBSECMBhavcopy()");
                string NameFromXML = ((string)xmlDoc.Element("BOD-Utility").Element("CM").Element("BSECM_BHAVCOPY").Element("NAME")).Trim();
                string DownloadFromLocal = ((string)xmlDoc.Element("BOD-Utility").Element("CM").Element("BSECM_BHAVCOPY").Element("LOCAL")).Trim();
                string[] arr_LocalFile = DownloadFromLocal.Split(',');
                AddToList($"BSECM Bhavcopy downloading.");

                bool filedownloaded = false;

                string[] arr_OldBhavcopyFiles;

                for (int i = 1; i < arr_BhavcopyFolderPath.Length; i++)
                {
                    if (!Directory.Exists(arr_BhavcopyFolderPath[i]))
                        Directory.CreateDirectory(arr_BhavcopyFolderPath[i]);
                    else
                    {
                        arr_OldBhavcopyFiles = Directory.GetFiles(arr_BhavcopyFolderPath[i], @"BhavCopy_BSE_CM_*.csv");
                        for (int j = 0; j < arr_OldBhavcopyFiles.Count(); j++)
                            File.Delete(arr_OldBhavcopyFiles[j]);
                    }
                }

                //var BhavcopyFileName = $"BSE_EQ_BHAVCOPY_{dateEdit_DownloadDate.DateTime.STR("ddMMyyyy")}.ZIP";

                //Added by Musharraf to check previous 7 days files
                var dateToCheck = dateEdit_DownloadDate.DateTime;
                var BhavcopyFileName = NameFromXML.Replace("$date:ddMMyyyy$", dateToCheck.ToString("ddMMyyyy")); /*ds_Config.GET("FILENAME", "BSECM_BHAVCOPY").ToString().Replace("$date:ddMMyyyy$", dateToCheck.ToString("ddMMyyyy"));*/
                // Added by Musharraf 6th April 2023
                // check if today is a weekend day
                //if (dateToCheck.DayOfWeek == DayOfWeek.Saturday || dateToCheck.DayOfWeek == DayOfWeek.Sunday)
                //{
                //    // if so, set the date to the previous Friday
                //    dateToCheck = dateToCheck.AddDays(-(int)dateToCheck.DayOfWeek).AddDays(-1);
                //}
                // check for the file in the previous 7 working days
                for (int i = 0; i < 7; i++)
                {

                    //BhavcopyFileName = $"BSE_EQ_BHAVCOPY_{dateToCheck.STR("ddMMyyyy")}.ZIP";
                    BhavcopyFileName = NameFromXML.Replace("$date:ddMMyyyy$", dateToCheck.ToString("ddMMyyyy"));
                    // download the file if it exists

                    try
                    {
                        using (WebClient webClient = new WebClient())
                        {
                            webClient.DownloadFile(BhavcopyURL[0] + BhavcopyFileName, arr_BhavcopyFolderPath[0] + BhavcopyFileName);
                        }

                        //arr_OldBhavcopyFiles = Directory.GetFiles(arr_BhavcopyFolderPath[0], @"BSE_EQ_BHAVCOPY_*.csv");
                        //for (int j = 0; j < arr_OldBhavcopyFiles.Count(); j++)
                        //    File.Delete(arr_OldBhavcopyFiles[j]);

                        using (ZipFile zip = ZipFile.Read(arr_BhavcopyFolderPath[0] + BhavcopyFileName))
                            zip.ExtractAll(arr_BhavcopyFolderPath[0], ExtractExistingFileAction.DoNotOverwrite);

                        File.Delete(arr_BhavcopyFolderPath[0] + BhavcopyFileName);

                        filedownloaded = true;
                        if (filedownloaded == true)
                        {
                            AddToList($"BSECM Bhavcopy: {BhavcopyFileName.Replace(".zip", ".csv")} downloaded successfully.");
                            _BSE_EQ_BHAVCOPY = BhavcopyFileName;
                            break;
                        }
                    }
                    catch (Exception ee) { _logger.Error(ee,"Download BSECM Bhavcopy from Web"); }
                    // exit the loop if the file is downloaded successfully

                    try
                    {
                        if (!filedownloaded)
                        {
                            //BhavCopy_BSE_CM_0_0_0_20231124_F_0000
                            BhavcopyFileName = $"BhavCopy_BSE_CM_0_0_0_{dateToCheck.STR("ddMMyyyy")}_F_0000.csv";
                            File.Copy(arr_LocalFile[0] + BhavcopyFileName, arr_BhavcopyFolderPath[0] + BhavcopyFileName, true);

                            filedownloaded = true;
                            AddToList($"BSECM Bhavcopy: {BhavcopyFileName.Replace(".zip", ".csv")} downloaded successfully.");
                            _BSE_EQ_BHAVCOPY = BhavcopyFileName;
                            break;
                        }
                    }
                    catch (Exception ee) { _logger.Error(ee,"Copying BSECM Bhavcopy from Local Folder"); }


                    // subtract a day from the date to check the previous day
                    dateToCheck = dateToCheck.AddDays(-1);

                    // skip weekends
                    if (dateToCheck.DayOfWeek == DayOfWeek.Saturday)
                    {
                        dateToCheck = dateToCheck.AddDays(-1);
                    }
                    else if (dateToCheck.DayOfWeek == DayOfWeek.Sunday)
                    {
                        dateToCheck = dateToCheck.AddDays(-2);
                    }
                }

                //end of the previous 7 days file check


                #region Copy from local File(shifted inside for loop)
                //try
                //{
                //    if (!filedownloaded)
                //    {
                //        BhavcopyFileName = $"BSE_EQ_BHAVCOPY_{dateToCheck.STR("ddMMyy")}.csv";
                //        File.Copy(BhavcopyURL[1] + BhavcopyFileName, arr_BhavcopyFolderPath[0] + BhavcopyFileName, true);

                //        filedownloaded = true;

                //    }
                //}
                //catch (Exception ee) { _logger.Error(ee); }
                #endregion
                if (filedownloaded)
                {
                    if (arr_BhavcopyFolderPath.Length > 1)
                    {
                        var arr_BhavcopyCSV = Directory.GetFiles(arr_BhavcopyFolderPath[0], "BhavCopy_BSE_EQ_*.csv")
                            .OrderByDescending(f => new FileInfo(f).CreationTime)
                            .ToArray();//modified by Musharraf changed EQ_ISINCODE_ to BSE_EQ_Bhavcopy
                        BhavcopyFileName = $"BhavCopy_BSE_CM_0_0_0_{dateToCheck.STR("ddMMyyyy")}_F_0000.csv"; ;

                        for (int i = 1; i < arr_BhavcopyFolderPath.Length; i++)
                        {
                            var oldcsvFiles = Directory.GetFiles(arr_BhavcopyFolderPath[i], "BSE_EQ_BHAVCOPY_*.csv");
                            for (int j = 0; j < oldcsvFiles.Count(); j++)
                                File.Delete(oldcsvFiles[j]);

                            if (arr_BhavcopyCSV.Length > 0 && File.Exists(arr_BhavcopyCSV[0]))
                            {
                                File.Copy(arr_BhavcopyCSV[0], arr_BhavcopyFolderPath[i] + BhavcopyFileName, true);
                            }
                        }
                    }

                    _logger.Debug($"BSECM Bhavcopy downloaded successfully in all Save-Paths");
                }
                else
                {
                    AddToList($"Unable to download BSECM Bhavcopy file.", true);
                }
            }
            catch (Exception ee) { _logger.Error(ee, $"Download BSECM Bhavcopy"); AddToList($"Unable to download BSECM Bhavcopy file.", true); }
        }

        private void InvokeDownloader(string[] SpanPath, string[] ExposurePath, string VaRExposurePath)
        {
            _logger.Debug("Executing InvokeDownloader():");
            try
            {
                object tempObj = null;
                ElapsedEventArgs tempE = null;

                //added on 28DEC2020 by Amey
                DownloadOTMScripFile(SpanPath);
                Thread.Sleep(1000);

                //Added by Snehadri on 11NOV2022
                DownloadNiftyExposureFile(ApplicationPath, ExposurePath);
                Thread.Sleep(1000);

                //Added by Snehadri on 11NOV20222
                CombineOTMFiles(SpanPath, ExposurePath);
                Thread.Sleep(1000);

                DownloadExposureFile(ApplicationPath, ExposurePath);
                Thread.Sleep(1000);

                DownloadBSEExposureFile(ApplicationPath, ExposurePath);
                Thread.Sleep(1000);

                //Added by Akshay on 13-10-2021 for Downloading CD Span
                DownloadCDExposureFile(ApplicationPath, ExposurePath);
                Thread.Sleep(1000);

                AddToList($"VaR Exposure file downloading.");
                AddToList($"Span file downloading.");
                AddToList($"CD Span file downloading.");    //Added by Akshay on 13-10-2021 for Downloading CD Span

                //DownloadSpan(tempObj, tempE, SpanPath);//Added for Testing
                //DownloadCDSpan(tempObj, tempE, SpanPath);//Added for Testing
                //DownloadVARExposure(tempObj, tempE, VaRExposurePath);//Added for Testing
                Parallel.Invoke(() => DownloadSpan(tempObj, tempE, SpanPath), () => DownloadCDSpan(tempObj, tempE, SpanPath), () => DownloadVARExposure(tempObj, tempE, VaRExposurePath), () => DownloadBSESpan(tempObj, tempE, SpanPath));

                if (this.InvokeRequired)
                    this.Invoke((MethodInvoker)(() => btn_DownloadSpan.Enabled = true));

                var timer = new System.Timers.Timer();
                timer.Interval = Convert.ToInt32(ds_Config.GET("INTERVAL", "SPAN-RECHECK-SECONDS")) * 1000;

                timer.Elapsed += (sender, e) => { Parallel.Invoke(() => DownloadSpan(tempObj, tempE, SpanPath), () => DownloadCDSpan(tempObj, tempE, SpanPath), () => DownloadVARExposure(tempObj, tempE, VaRExposurePath), () => DownloadBSESpan(tempObj, tempE, SpanPath)); };
                timer.AutoReset = true;
                timer.Enabled = true;
            }
            catch (Exception ee) { _logger.Error(ee, "In InvokeDownloader(): "); }
        }

        private void DownloadExposureFile(string IndexFolderPath, string[] ExposureFolderPath)
        {
            try
            {
                bool filedownloaded = false;
                string[] ExposureFileUrl = ds_Config.GET("URLs", "EXPOSURE").SPL(',');
                string ExposureFileName = "ael_" + DateTime.Now.STR("ddMMyyyy");
                string ExposureFilePath = ExposureFolderPath[0] + ExposureFileName + ".csv";

                AddToList($"Exposure file [{ExposureFileName}] downloading.");

                //using API 
                try
                {
                    var response = nNSEUtils.Instance.DownloadCommonFile(en_FolderTypes.FO_AEL, ExposureFileName + ".csv", ExposureFolderPath[0]);
                    _logger.Debug("DownloadExposureFile " + ExposureFileName + " | API Response: " + JsonConvert.SerializeObject(response.Response.ResponseStatus));
                    if (response.ResponseStatus == en_ResponseStatus.SUCCESS)
                    {
                        //DecompressGZAndDelete(new FileInfo(SpanFilePath + @"\" + OTMFileName), "");
                        //_logger.Debug("Extracted OTM : " + true);
                        filedownloaded = true;
                    }
                }
                catch (Exception ee)
                {
                    _logger.Error(ee,"Downloading Exposure file using API");
                }

                // using URL
                try
                {
                    if (!filedownloaded)
                    {
                        string url = ExposureFileUrl[1] + "/ael_" + DateTime.Now.STR("ddMMyyyy") + ".csv";

                        using (WebClient webClient = new WebClient())
                        {
                            webClient.DownloadFile(url, ExposureFilePath);
                        }
                        filedownloaded = true;
                    }
                }
                catch (Exception ee) { _logger.Error(ee, "Downloading Exposure file using Web"); }

                #region FTP decommissioned
                // using FTP
                //if (!filedownloaded)
                //{
                //    try
                //    {
                //        using (WebClient webClient = new WebClient())
                //        {
                //            //Added to login and download from NSE FTP link. 16MAR2021-Amey
                //            webClient.Credentials = new NetworkCredential(dict_FTPCred["FO"].Username, dict_FTPCred["FO"].Password);

                //            webClient.DownloadFile(ExposureFileUrl[0] + ExposureFileName + ".csv", ExposureFilePath);
                //            filedownloaded = true;
                //        }
                //    }
                //    catch (Exception ee) { _logger.Error(ee); }

                //}
                #endregion
                // using File
                try
                {
                    if (!filedownloaded)
                    {
                        string filename = ExposureFileUrl[2] + ExposureFileName + ".csv";
                        File.Copy(filename, ExposureFilePath, true);
                        filedownloaded = true;
                    }
                }
                catch (Exception ee) { _logger.Error(ee, "Copying Exposure file using Local Folder"); }

                if (filedownloaded)
                {
                    _logger.Debug("FOExposure FileName : " + ExposureFileName);
                    AddIndexExposure(ExposureFilePath, IndexFolderPath + @"\IndexExposure.csv");

                    for (int i = 1; i < ExposureFolderPath.Length; i++)
                    {
                        File.Copy(ExposureFolderPath[0] + ExposureFileName + ".csv", ExposureFolderPath[i] + ExposureFileName + ".csv", true);
                    }

                    AddToList($"Exposure file [{ExposureFileName}] downloaded successfully.");
                }
                else
                {
                    AddToList($"Exposure file failed to download.", true);
                }
            }
            catch (Exception expEX)
            {
                _logger.Error(expEX, "DownloadExposureFile");

                AddToList($"Exposure file downloaded failed.", true);
            }
        }

        private void AddIndexExposure(string ExposureFilePath, string IndexExposureFilePath)
        {
            try
            {
                using (FileStream stream = File.Open(ExposureFilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    using (StreamReader reader = new StreamReader(stream))
                    {
                        string _line = reader.ReadToEnd();
                        StreamWriter _streamWriter = File.AppendText(ExposureFilePath);

                        using (FileStream _IndexStream = File.Open(IndexExposureFilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                        {
                            using (StreamReader _IndexReader = new StreamReader(_IndexStream))
                            {
                                string _innerVal;
                                while ((_innerVal = _IndexReader.ReadLine()) != null)
                                {
                                    string[] _strSep = _innerVal.Split(',');
                                    if (_line.Contains("," + _strSep[0] + ",")) continue;
                                    if (_strSep.Count() >= 4)
                                    {
                                        string LineToWrite = $"1,{_strSep[0]},{_strSep[1]},{_strSep[2]},{_strSep[3]},{_strSep[4]} {Environment.NewLine}";
                                        _streamWriter.WriteAsync(LineToWrite);
                                        _streamWriter.Flush();
                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception err)
            {
                _logger.Error(err, "AddIndexExposure");
            }
        }

        private void DownloadSpan(Object source, ElapsedEventArgs e, string[] SpanFilePath)
        {
            try
            {
                if (this.InvokeRequired)
                    this.Invoke((MethodInvoker)(() => btn_DownloadSpan.Enabled = false));

                IsSpanFileDownloading = true;
                _logger.Debug("Outside IF Index : " + SpanIndex + " SpanFileExtensions Count : " + arr_SpanFileExtensions.Count());
                string[] SpanFileConfigURL = ds_Config.GET("URLs", "SPAN").SPL(',');
                if (SpanIndex < arr_SpanFileExtensions.Count())
                {
                    SpanFileDownloaded = false;

                    var SpanFileExactName = "nsccl." + DateTime.Now.Year.STR("0000") + DateTime.Now.Month.STR("00") + DateTime.Now.Date.STR("dd") + ".";
                    SpanFileExactName += arr_SpanFileExtensions[SpanIndex] + ".spn.gz";

                    var SpanFileName = SpanFileConfigURL[0] + SpanFileExactName;

                    _logger.Debug("Attempted SpanFileName : " + SpanFileName);

                    if (!Directory.Exists(SpanFilePath[0] + "TEMP"))
                        Directory.CreateDirectory(SpanFilePath[0] + "TEMP");

                    var TempSpanFilepPath = SpanFilePath[0] + "TEMP\\";

                    if (!File.Exists(SpanFilePath[0] + $"nsccl.{DateTime.Today.ToString("yyyyMMdd")}." + arr_SpanFileExtensions[SpanIndex] + ".spn"))   //changes by nikhil 
                    {
                        var DecmpressSpanFileName = $"nsccl.{DateTime.Today.ToString("yyyyMMdd")}." + arr_SpanFileExtensions[SpanIndex] + ".spn";

                        try
                        {
                            var response = nNSEUtils.Instance.DownloadCommonFile(en_FolderTypes.FO_SPAN, SpanFileExactName, TempSpanFilepPath);
                            _logger.Debug("Download span  " + SpanFileExactName + " | API RESPONSE :  " + JsonConvert.SerializeObject(response.Response.ResponseStatus));
                            if (response.ResponseStatus == en_ResponseStatus.SUCCESS)
                            {
                                var decompressfilePath = DecompressGZAndDelete(new FileInfo(TempSpanFilepPath + SpanFileExactName));
                                if (decompressfilePath != "")
                                {
                                    SpanFileDownloaded = true;
                                }

                            }
                        }
                        catch (Exception EE)
                        {
                            _logger.Error(EE);
                        }

                        try
                        {
                            if (!SpanFileDownloaded)
                            {
                                string extension = string.Empty;
                                if (arr_SpanFileExtensions[SpanIndex].Length > 1)
                                {
                                    extension = arr_SpanFileExtensions[SpanIndex].SUB(0, 1) + arr_SpanFileExtensions[SpanIndex].SUB(2, 1);
                                }
                                else
                                {
                                    extension = arr_SpanFileExtensions[SpanIndex];
                                }

                                string url = SpanFileConfigURL[1] + $"nsccl.{DateTime.Today.ToString("yyyyMMdd")}." + extension + ".zip";
                                var TimeToWait = int.Parse((xmlDoc.Element("BOD-Utility").Element("INTERVAL").Element("SPAN-WAIT-SECONDS")).Value);//Added by musharraf
                                                                                                                                                   //var Seconds = TimeSpan.FromSeconds(TimeToWait);//Added by Musharraf

                                try
                                {

                                    WebClient client = new WebClient();
                                    Thread th = new Thread(() => client.DownloadFileTaskAsync(new Uri(url), TempSpanFilepPath + SpanFileExactName));
                                    th.Start();
                                    _logger.Debug("Span Thread init");
                                    Thread.Sleep(TimeToWait * 1000);
                                    client.CancelAsync();
                                    th.Abort();

                                } catch (Exception EE) { _logger.Debug(EE.Message + "File Couldn't be Downloaded"); }

                                //if condition Added by Musharraf 
                                if (File.Exists(TempSpanFilepPath + SpanFileExactName))
                                {
                                    using (ZipFile zip = ZipFile.Read(TempSpanFilepPath + SpanFileExactName))
                                        zip.ExtractAll(TempSpanFilepPath, ExtractExistingFileAction.DoNotOverwrite);

                                    _logger.Debug($"Span File decompressed: {TempSpanFilepPath + SpanFileExactName}");
                                    File.Delete(TempSpanFilepPath + SpanFileExactName);
                                    SpanFileDownloaded = true;
                                }
                            }
                        }
                        catch (Exception ee) { _logger.Error(ee); }

                        #region FTP decommissioned
                        //FTP
                        //if (!SpanFileDownloaded)
                        //{
                        //    Thread th_FTP = new Thread(() => DownloadviaFTP(SpanFileName, TempSpanFilepPath, SpanFilePath, SpanFileExactName));
                        //    th_FTP.Start();
                        //    Thread.Sleep(SpanWaitSeconds);
                        //    th_FTP.Abort();
                        //}
                        #endregion
                        try
                        {
                            if (!SpanFileDownloaded)
                            {
                                string filename = SpanFileConfigURL[2] + $"nsccl.{DateTime.Today.ToString("yyyyMMdd")}." + arr_SpanFileExtensions[SpanIndex] + ".spn";
                                File.Copy(filename, TempSpanFilepPath + $"nsccl.{DateTime.Today.ToString("yyyyMMdd")}." + arr_SpanFileExtensions[SpanIndex] + ".spn", true);
                                SpanFileDownloaded = true;
                                _logger.Debug($"Copying from Local: Span file [{SpanFileExactName}] downloaded successfully.");
                            }
                        }
                        catch (Exception ee) { _logger.Error(ee, "Copying from Local:"); }

                    }

                    if (SpanFileDownloaded)
                    {
                        var isCurrptFile = false;
                        try
                        {
                            _logger.Debug($"Span file [{SpanFileExactName}] downloaded successfully.");
                            var decompressSpanFileName = $"nsccl.{DateTime.Today.ToString("yyyyMMdd")}." + arr_SpanFileExtensions[SpanIndex] + ".spn";
                            var decompressfilePath = TempSpanFilepPath + decompressSpanFileName;
                            var tryOpen = File.ReadAllText(decompressfilePath);
                            tryOpen = "";

                            XDocument doc = new XDocument();
                            doc = XDocument.Load(decompressfilePath);

                            if (!File.Exists(SpanFilePath[0] + decompressSpanFileName))
                                File.Copy(decompressfilePath, SpanFilePath[0] + decompressSpanFileName);

                            File.Delete(decompressfilePath);

                        }
                        catch (Exception ee)
                        {
                            isCurrptFile = true;
                            _logger.Error(ee);
                        }

                        if (!isCurrptFile)
                        {
                            AddToList($"Span file [{SpanFileExactName}] downloaded successfully.");
                            _logger.Debug("Downloaded SpanFileName : " + SpanFileName);

                            string filename = $"nsccl.{DateTime.Today.ToString("yyyyMMdd")}." + arr_SpanFileExtensions[SpanIndex] + ".spn";
                            for (int i = 1; i < SpanFilePath.Length; i++)
                                File.Copy(SpanFilePath[0] + filename, SpanFilePath[i] + filename, true);

                            try
                            {
                                arr_SpanFileExtensions = arr_SpanFileExtensions.Take(arr_SpanFileExtensions.Count() - (arr_SpanFileExtensions.Count() - SpanIndex)).ToArray();
                                SpanIndex = 0;
                                _logger.Debug("After Slice Index : " + SpanIndex + " FOSpanFileExtensions Count : " + arr_SpanFileExtensions.Count());
                            }
                            catch (Exception ee) { _logger.Error(ee, "Slicing Array"); }

                            _logger.Debug("--------------------------------------------------------------------------------");
                        }
                        else
                        {
                            _logger.Debug("Download Failed SpanFileName : " + SpanFileName);
                            _logger.Debug("Currpt File downloaded");
                        }
                    }
                    else
                    {
                        _logger.Debug("Download Failed SpanFileName : " + SpanFileName);

                        SpanIndex++;
                        if (SpanIndex < arr_SpanFileExtensions.Count())
                            DownloadSpan(source, e, SpanFilePath);
                        else
                            SpanIndex = 0;
                    }
                }

                IsSpanFileDownloading = false;
                if (this.InvokeRequired)
                    this.Invoke((MethodInvoker)(() => btn_DownloadSpan.Enabled = true));
            }
            catch (Exception ee)
            {
                IsSpanFileDownloading = false;
                SpanIndex = 0;
                _logger.Error(ee, "DownloadFOSpan");
                if (this.InvokeRequired)
                    this.Invoke((MethodInvoker)(() => btn_DownloadSpan.Enabled = true));
            }
        }

        private void DownloadviaFTP(string SpanFileName, string TempSpanFilePath, string[] SpanFilePath, string SpanFileExactName)
        {

            try
            {
                try
                {
                    using (WebClient webClient = new WebClient())
                    {

                        //Added to login and download from NSE FTP link. 16MAR2021-Amey
                        webClient.Credentials = new NetworkCredential(dict_FTPCred["FO"].Username, dict_FTPCred["FO"].Password);

                        webClient.DownloadFile(SpanFileName, TempSpanFilePath + SpanFileExactName);
                    }

                    var decompressfilePath = DecompressGZAndDelete(new FileInfo(TempSpanFilePath + SpanFileExactName));
                    SpanFileDownloaded = true;

                }
                catch (Exception EE)
                {
                    _logger.Error(EE);
                }

            }
            catch (Exception ee) { _logger.Error(ee); }
        }
        //Added by Akshay on 13-10-2021 for Downloading CD Span
        private void DownloadCDSpan(Object source, ElapsedEventArgs e, string[] SpanFilePath)
        {
            try
            {
                if (SpanFilePath.Length > 0)
                {
                    if (!Directory.Exists(SpanFilePath[0] + "TEMP"))
                        Directory.CreateDirectory(SpanFilePath[0] + "TEMP");

                    var TempSpanFilePath = SpanFilePath[0] + "TEMP\\";

                    _logger.Debug("Outside IF Index : " + CDSpanIndex + " SpanFileExtensions Count : " + arr_CDSpanFileExtensions.Count());
                    var FileDownloaded = false;

                    if (CDSpanIndex < arr_CDSpanFileExtensions.Count())
                    {
                        var SpanFileConfigURL = ds_Config.GET("URLs", "CD-SPAN");
                        var SpanFileExactName = "nsccl_ix." + DateTime.Now.Year.STR("0000") + DateTime.Now.Month.STR("00") + DateTime.Now.Date.STR("dd") + ".";
                        SpanFileExactName += arr_CDSpanFileExtensions[CDSpanIndex] + ".spn.gz";

                        var CDSpanFileName = SpanFileConfigURL + SpanFileExactName;

                        _logger.Debug("CDSpanFileName Before : " + CDSpanFileName);

                        var SpanFileNme = "";
                        var DownlodedFromFTP = false;

                        try
                        {
                            var response = nNSEUtils.Instance.DownloadCommonFile(en_FolderTypes.CD_SPAN, SpanFileExactName, TempSpanFilePath);
                            _logger.Debug("Download CDSspan  " + SpanFileExactName + " | API RESPONSE :  " + JsonConvert.SerializeObject(response.Response.ResponseStatus));
                            if (response.ResponseStatus == en_ResponseStatus.SUCCESS)
                            {
                                SpanFileNme = DecompressGZAndDelete(new FileInfo(TempSpanFilePath + SpanFileExactName));
                            }
                        }
                        catch (Exception ee)
                        {
                            _logger.Error(ee, $"DownloadCDSpan {SpanFileExactName}");
                        }

                        #region FTP
                        //if (SpanFileNme == "")
                        //{
                        //    try
                        //    {
                        //        using (WebClient webClient = new WebClient())
                        //        {
                        //            //Added to login and download from NSE FTP link. 16MAR2021-Amey
                        //            webClient.Credentials = new NetworkCredential(dict_FTPCred["CD"].Username, dict_FTPCred["CD"].Password);

                        //            webClient.DownloadFile(CDSpanFileName, TempSpanFilePath + SpanFileExactName);
                        //        }
                        //        SpanFileNme = DecompressGZAndDelete(new FileInfo(TempSpanFilePath + SpanFileExactName));

                        //    }
                        //    catch (Exception ee)
                        //    {
                        //        _logger.Error(ee, $"DownloadCDSpan {SpanFileExactName}");

                        //    }
                        //}
                        #endregion
                        _logger.Debug("CDSpanFileName After : " + CDSpanFileName);
                        // var SpanFileNme = DecompressGZAndDelete(new FileInfo(TempSpanFilePath + SpanFileExactName));
                        try
                        { //download from local file
                            if (!FileDownloaded)
                            {
                               
                                string sourcefile = SpanFileConfigURL + "nsccl_ix." + DateTime.Now.Year.STR("0000") + DateTime.Now.Month.STR("00") + DateTime.Now.Date.STR("dd") + "." + arr_CDSpanFileExtensions[CDSpanIndex] + ".spn";
                                string destination = TempSpanFilePath + "nsccl_ix." + DateTime.Now.Year.STR("0000") + DateTime.Now.Month.STR("00") + DateTime.Now.Date.STR("dd") + "." + arr_CDSpanFileExtensions[CDSpanIndex] + ".spn";

                                _logger.Debug("Copying CD Span file from local drive Source : "+ sourcefile +" | Destination : "+ destination);

                                File.Copy(sourcefile, destination, true);
                                
                                SpanFileNme = destination;
                                _logger.Debug($"Copying from Local: Span file [{Path.GetFileName(destination)}] downloaded successfully.");
                            }
                        }
                        catch (Exception ee) { _logger.Error(ee, "Exception Copying from Local CDspan:"); }
                        if (SpanFileNme != "")
                        {
                            //trying to open
                            try
                            {
                                var tryOpen = File.ReadAllText(SpanFileNme);
                                tryOpen = "";

                                XDocument doc = new XDocument();
                                doc = XDocument.Load(SpanFileNme);

                                var decompressFileName = "nsccl_ix." + DateTime.Now.Year.STR("0000") + DateTime.Now.Month.STR("00") + DateTime.Now.Date.STR("dd") + "." + arr_CDSpanFileExtensions[CDSpanIndex] + ".spn";
                                if (!File.Exists(SpanFilePath[0] + decompressFileName))
                                    File.Copy(SpanFileNme, SpanFilePath[0] + decompressFileName);

                                File.Delete(SpanFileNme);
                                FileDownloaded = true;
                                AddToList($"CD Span file [{SpanFileExactName}] downloaded successfully.");
                            }
                            catch (Exception ee)
                            {
                                _logger.Debug("Currpt file downloaded or File opened somewhere");
                                _logger.Error(ee);
                                FileDownloaded = false;
                            }

                        }

                        if (FileDownloaded)
                        {

                            string filename = $"nsccl_ix.{DateTime.Today.ToString("yyyyMMdd")}." + arr_CDSpanFileExtensions[CDSpanIndex] + ".spn";

                            for (int i = 1; i < SpanFilePath.Length; i++)
                                File.Copy(SpanFilePath[0] + filename, SpanFilePath[i] + filename, true);

                            _logger.Debug($"CD Span file [{SpanFileExactName}] downloaded successfully.");

                            _logger.Debug("Extracted : " + true);

                            try
                            {
                                arr_CDSpanFileExtensions = arr_CDSpanFileExtensions.Take(arr_CDSpanFileExtensions.Count() - (arr_CDSpanFileExtensions.Count() - CDSpanIndex)).ToArray();
                                CDSpanIndex = 0;
                                _logger.Debug("After Slice Index : " + CDSpanIndex + " CDSpanFileExtensions Count : " + arr_CDSpanFileExtensions.Count());
                            }
                            catch (Exception ee) { _logger.Error(ee, "Slicing Array"); }

                        }
                        else
                        {
                            CDSpanIndex++;
                            if (CDSpanIndex < arr_CDSpanFileExtensions.Count())
                                DownloadCDSpan(source, e, SpanFilePath);
                            else
                                CDSpanIndex = 0;
                        }

                    }
                }
            }
            catch (Exception ee)
            {
                CDSpanIndex = 0;
                _logger.Error(ee, "DownloadCDSpan");
            }
        }

        private void DownloadOTMScripFile(string[] SpanFilePath)
        {
            try
            {
                string[] OTMFileURL = ds_Config.Tables["URLs"].Rows[0]["OTM"].ToString().SPL(',');
                string OTMFileName = $"F_AEL_OTM_CONTRACTS_{DateTime.Now.ToString("ddMMyyyy")}.csv.gz";
                bool filedownloaded = false;

                AddToList($"OTM Exposure file [{OTMFileName}] downloading.");

                try
                {
                    var response = nNSEUtils.Instance.DownloadCommonFile(en_FolderTypes.FO_OTM, OTMFileName, SpanFilePath[0]);
                    _logger.Debug("DownloadOTMScripFile " + OTMFileName + " | API Response: " + JsonConvert.SerializeObject(response));
                    if (response.ResponseStatus == en_ResponseStatus.SUCCESS)
                    {
                        DecompressGZAndDelete(new FileInfo(SpanFilePath[0] + OTMFileName), "");
                        _logger.Debug("Extracted OTM : " + true);
                        filedownloaded = true;
                    }
                }
                catch (Exception ee)
                {
                    _logger.Error(ee);
                }

                #region FTP
                //if (!filedownloaded)
                //{
                //    try
                //    {
                //        using (WebClient webClient = new WebClient())
                //        {
                //            //Added to login and download from NSE FTP link. 16MAR2021-Amey
                //            webClient.Credentials = new NetworkCredential(dict_FTPCred["FO"].Username, dict_FTPCred["FO"].Password);
                //            webClient.DownloadFile(OTMFileURL[0] + OTMFileName, SpanFilePath[0] + OTMFileName);
                //        }

                //        DecompressGZAndDelete(new FileInfo(SpanFilePath[0] +  OTMFileName), "");

                //        _logger.Debug("Extracted OTM : " + true);
                //        filedownloaded = true;

                //    }
                //    catch (Exception ee) { _logger.Error(ee); }
                //}
                #endregion
                if (!filedownloaded)
                {
                    //Added by Musharraf on 08-05-2023 download from local Files
                    try
                    {
                        string OTMFileNameUnzip = $"F_AEL_OTM_CONTRACTS_{DateTime.Now.ToString("ddMMyyyy")}.csv";
                        string filename = OTMFileURL[1] + OTMFileNameUnzip;
                        string destination = SpanFilePath[0] + OTMFileNameUnzip;
                        File.Copy(filename, destination, true);
                        //DecompressGZAndDelete(new FileInfo(SpanFilePath[0] + @"\" + OTMFileNameUnzip), "");
                        _logger.Debug("Extracted OTM : " + true);
                        filedownloaded = true;
                    }
                    catch (Exception ee)
                    {
                        _logger.Error(ee);
                    }
                }

                if (filedownloaded)
                {
                    var FileName = $"F_AEL_OTM_CONTRACTS_{DateTime.Now.ToString("ddMMyyyy")}.csv";

                    for (int i = 1; i < SpanFilePath.Length; i++)
                    {
                        File.Copy(SpanFilePath[0] + FileName, SpanFilePath[i] + FileName, true);
                    }

                    AddToList($"OTM file [{OTMFileName}] downloaded successfully.");

                }
                else
                    AddToList($"Unable to download OTM file.", true);
            }
            catch (Exception ee) { _logger.Error(ee); AddToList($"Unable to download OTM file.", true); }
        }

        // Added by Snehadri on 11NOV2022
        private void DownloadNiftyExposureFile(string IndexFolderPath, string[] ExposureFolderPath)
        {
            try
            {
                bool filedownloaded = false;
                string[] ExposureFileUrl = ds_Config.GET("URLs", "NIFTY_EXPOSURE").SPL(',');
                string ExposureFileName = "ael_NIFTY_Options";
                string ExposureFilePath = ExposureFolderPath[0] + ExposureFileName + ".csv";

                AddToList($"Nifty Exposure file [{ExposureFileName}] downloading.");

                //using API 
                try
                {
                    var response = nNSEUtils.Instance.DownloadCommonFile(en_FolderTypes.FO_NIFTY_AEL, ExposureFileName + ".csv", ExposureFolderPath[0]);
                    _logger.Debug("DownloadNiftyExposureFile " + ExposureFileName + " | API Response: " + JsonConvert.SerializeObject(response));
                    if (response.ResponseStatus == en_ResponseStatus.SUCCESS)
                    {
                        //DecompressGZAndDelete(new FileInfo(SpanFilePath + @"\" + OTMFileName), "");
                        //_logger.Debug("Extracted OTM : " + true);
                        filedownloaded = true;
                    }
                }
                catch (Exception ee)
                {
                    _logger.Error(ee);
                }
                #region FTP
                // using FTP
                //if (!filedownloaded)
                //{
                //    try
                //    {
                //        using (WebClient webClient = new WebClient())
                //        {
                //            //Added to login and download from NSE FTP link. 16MAR2021-Amey
                //            webClient.Credentials = new NetworkCredential(dict_FTPCred["FO"].Username, dict_FTPCred["FO"].Password);

                //            webClient.DownloadFile(ExposureFileUrl[0] + ExposureFileName + ".csv", ExposureFilePath);
                //            filedownloaded = true;
                //        }
                //    }
                //    catch (Exception ee) { _logger.Error(ee); }
                //}
                #endregion
                // using File
                try
                {
                    if (!filedownloaded)
                    {   //Added by Musharraf on 10th June23
                        string[] nonNullElements = ExposureFileUrl.Where(url => !string.IsNullOrWhiteSpace(url)).ToArray();
                        string filename = nonNullElements.ElementAtOrDefault(1);

                        filename += ExposureFileName + ".csv";
                        File.Copy(filename, ExposureFilePath, true);
                        filedownloaded = true;

                    }
                }
                catch (Exception ee) { _logger.Error(ee); }

                if (filedownloaded)
                {

                    for (int i = 1; i < ExposureFolderPath.Length; i++)
                    {
                        File.Copy(ExposureFolderPath[0] + ExposureFileName + ".csv", ExposureFolderPath[i] + ExposureFileName + ".csv", true);
                    }

                    _logger.Debug("Nifty Exposure FileName : " + ExposureFileName);
                    AddToList($"Nifty Exposure file [{ExposureFileName}] downloaded successfully.");
                }
                else
                {
                    AddToList($"Nifty Exposure file failed to download.", true);
                }
            }
            catch (Exception expEX)
            {
                _logger.Error(expEX, "DownloadNiftyExposureFile ");
            }
        }

        // Added by Snehadri on 11NOV2022
        private void CombineOTMFiles(string[] SpanPath, string[] ExposurePath)
        {
            try
            {
                var Directory = new DirectoryInfo(ExposurePath[0]);

                var OTMFile = Directory.GetFiles("F_AEL_OTM_CONTRACTS*.csv")
                               .OrderByDescending(f => f.LastWriteTime)
                               .FirstOrDefault();

                var NiftyOTMFile = Directory.GetFiles("ael_NIFTY_Options*.csv")
                           .OrderByDescending(f => f.LastWriteTime)
                           .FirstOrDefault();


                var BhavcopyDirectory = new DirectoryInfo(@"C:/Prime");

                var FOBhavcopy = BhavcopyDirectory.GetFiles("NSE_FO_bhavcopy*.csv")
                               .OrderByDescending(f => f.LastWriteTime)
                               .First();

                if (OTMFile != null && NiftyOTMFile != null)
                {
                    AddToList("Adding Scrips in OTM File");

                    var _Result = AddOTMExposure(FOBhavcopy.FullName, ExposurePath);

                    if (!_Result)
                    {
                        AddToList("Adding Scrips in OTM File Failed", true);
                        DownloadOTMScripFile(SpanPath);
                    }
                }
            }
            catch (Exception ee)
            {
                _logger.Error(ee, "CombineOTMFiles ");
            }
        }

        private bool AddOTMExposure(string FOBhavcopyPath, string[] ExposureFolder)
        {
            bool result = false;

            try
            {
                ConcurrentDictionary<string, OTMFileData> dict_OTMFileData = new ConcurrentDictionary<string, OTMFileData>();
                ConcurrentDictionary<string, NiftyOTMFile> dict_NiftyOTMFile = new ConcurrentDictionary<string, NiftyOTMFile>();

                var list_FOBhavcopy = Exchange.ReadFOBhavcopy(FOBhavcopyPath, true);

                var list_NiftyContracts = list_FOBhavcopy.Where(v => v.Symbol == "NIFTY").ToList();

                var OTMFileName = $"F_AEL_OTM_CONTRACTS_{DateTime.Now.ToString("ddMMyyyy")}.csv";
                var NiftyOTMFileName = "ael_NIFTY_Options.csv";

                using (FileStream fs = File.Open(ExposureFolder[0] + OTMFileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    using (BufferedStream bs = new BufferedStream(fs))
                    {
                        using (StreamReader sr = new StreamReader(bs))
                        {
                            string strData = string.Empty;

                            while ((strData = sr.ReadLine()) != null)
                            {
                                var data = strData.ToUpper().Split(',').Select(v => v.Trim()).ToArray();

                                OTMFileData oTMFileData = new OTMFileData()
                                {
                                    InstName = data[0],
                                    Symbol = data[1],
                                    ExpiryDate = DateTime.Parse(data[2]),
                                    StrikePrice = data[3],
                                    ScripType = data[4],
                                    Percentage = data[6]
                                };

                                string _Key = $"{oTMFileData.Symbol}|{oTMFileData.ScripType}|{oTMFileData.StrikePrice}|{oTMFileData.ExpiryDate.ToString("dd-MMM-yy")}";

                                dict_OTMFileData.TryAdd(_Key, oTMFileData);
                            }
                        }
                    }
                }

                using (FileStream fs = File.Open(ExposureFolder[0] + NiftyOTMFileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    using (BufferedStream bs = new BufferedStream(fs))
                    {
                        using (StreamReader sr = new StreamReader(bs))
                        {
                            string strData = string.Empty;

                            while ((strData = sr.ReadLine()) != null)
                            {
                                var data = strData.ToUpper().Split(',').Select(v => v.Trim()).ToArray();


                                var _Symbol = data[1];
                                var _Expiry = data[3];

                                string _Key = $"{_Symbol}|{_Expiry}";

                                if (dict_NiftyOTMFile.TryGetValue(_Key, out NiftyOTMFile oNiftyOTMFile))
                                {
                                    if (data[2] == "OTM")
                                        oNiftyOTMFile.OTMPercentage = data[4];
                                    else
                                        oNiftyOTMFile.OTHPercentage = data[4];
                                }
                                else
                                {
                                    NiftyOTMFile OTMFileData = new NiftyOTMFile()
                                    {
                                        Symbol = _Symbol,
                                        ExpiryDate = _Expiry
                                    };

                                    if (data[2] == "OTM")
                                    {
                                        OTMFileData.OTMPercentage = data[4];
                                        OTMFileData.OTHPercentage = "0";
                                    }
                                    else
                                    {
                                        OTMFileData.OTHPercentage = data[4];
                                        OTMFileData.OTMPercentage = "0";
                                    }

                                    dict_NiftyOTMFile.TryAdd(_Key, OTMFileData);

                                }
                            }
                        }
                    }
                }

                StringBuilder sb_NewData = new StringBuilder();

                var _Today = DateTime.Today.AddHours(15).AddMinutes(15);

                foreach (var _Contract in list_NiftyContracts)
                {
                    if (GetMonthDifference(_Contract.Expiry, _Today) < 9) continue;

                    string _Key = $"{_Contract.Symbol}|{_Contract.ScripType}|{_Contract.StrikePrice}|{_Contract.Expiry.ToString("dd-MMM-yy")}";

                    var _NKey = $"{_Contract.Symbol}|{_Contract.Expiry.ToString("dd-MMM-yy").ToUpper()}";

                    if (!dict_NiftyOTMFile.ContainsKey(_NKey)) continue;

                    var _NValue = dict_NiftyOTMFile[_NKey];

                    if (!dict_OTMFileData.ContainsKey(_Key))
                    {
                        var str = $"{_Contract.Instrument},{_Contract.Symbol},{_Contract.Expiry.ToString("dd-MMM-yyyy").ToUpper()},{_Contract.StrikePrice},{_Contract.ScripType},0,{_NValue.OTHPercentage}";
                        sb_NewData.AppendLine(str);
                    }
                }

                for (int i = 0; i < ExposureFolder.Length; i++)
                    File.AppendAllText(ExposureFolder[i] + OTMFileName, sb_NewData.ToString());

                AddToList("Added Scrips in OTM File");

                result = true;
            }
            catch (Exception ee)
            {
                _logger.Error(ee, "AddOTMExposure ");
            }

            return result;
        }

        public static int GetMonthDifference(DateTime startDate, DateTime endDate)
        {
            int monthsApart = (12 * (startDate.Year - endDate.Year)) + (startDate.Month - endDate.Month);
            return Math.Abs(monthsApart);
        }

        private void DownloadVARExposure(Object source, ElapsedEventArgs e, string VaRExposureFilePath)
        {
            try
            {
                _logger.Debug("[VaRExposure] Outside IF Index : " + VaRIndex + " VaRFileExtensions Count : " + arr_VaRFileExtensions.Count());
                if (VaRIndex < arr_VaRFileExtensions.Count())
                {
                    bool filedownloaded = false;

                    string[] VaRUrl = ds_Config.GET("URLs", "VAREXPOSURE").SPL(',');

                    string VaRFileName = VaRUrl[0];
                    VaRFileName += "C_VAR1_" + DateTime.Now.STR("ddMMyyyy") + "_";
                    VaRFileName += arr_VaRFileExtensions[VaRIndex] + ".DAT";

                    var ExactSpanName = VaRFileName.SUB((VaRFileName.CON("\\") ? VaRFileName.LastIndexOf("\\") : VaRFileName.LastIndexOf("/")) + 1);

                    _logger.Debug("[VaRExposure] VaRFileName Before : " + VaRFileName);

                    File.Delete(VaRExposureFilePath + ExactSpanName);

                    try
                    {
                        var response = nNSEUtils.Instance.DownloadCommonFile(en_FolderTypes.CM_VAREXPOSURE, ExactSpanName, VaRExposureFilePath);
                        _logger.Debug("DownloadVarExposureFile API Response: " + JsonConvert.SerializeObject(response));
                        filedownloaded = response.ResponseStatus == en_ResponseStatus.SUCCESS ? true : false;
                        if (File.Exists(VaRExposureFilePath + ExactSpanName))
                        {
                            var arr_bytes = File.ReadAllBytes(VaRExposureFilePath + ExactSpanName);
                            if (arr_bytes.Length > 0)
                            {
                                filedownloaded = true;
                                _logger.Debug($"DownloadVARExposure API: File {ExactSpanName}");
                            }
                            else
                            {
                                File.Delete(VaRExposureFilePath + ExactSpanName);
                                filedownloaded = false;
                                _logger.Debug($"DownloadVARExposure API: File Deleted {ExactSpanName} because Byteslength: {arr_bytes.Length} ");
                            }
                        }
                    }
                    catch (Exception ee) { _logger.Error(ee); }

                    // using Url
                    try
                    {
                        if (!filedownloaded)
                        {
                            string url = VaRUrl[1] + "C_VAR1_" + DateTime.Now.STR("ddMMyyyy") + "_" + arr_VaRFileExtensions[VaRIndex] + ".DAT";

                            var TimeToWait = int.Parse((xmlDoc.Element("BOD-Utility").Element("INTERVAL").Element("SPAN-WAIT-SECONDS")).Value);//Added by musharraf

                            WebClient client = new WebClient();
                            Thread th = new Thread(() => client.DownloadFileTaskAsync(new Uri(url), VaRExposureFilePath + ExactSpanName));
                            th.Start();
                            _logger.Debug("VAR Thread init");
                            Thread.Sleep(TimeToWait * 1000);
                            client.CancelAsync();
                            th.Abort();

                            _logger.Debug("Var thread status | " + th.IsAlive.ToString());

                            if (File.Exists(VaRExposureFilePath + ExactSpanName))
                            {
                                var arr_bytes = File.ReadAllBytes(VaRExposureFilePath + ExactSpanName);
                                if (arr_bytes.Length > 0)
                                    filedownloaded = true;
                                else
                                    File.Delete(VaRExposureFilePath + ExactSpanName);
                            }
                        }
                    }
                    catch (Exception ee) { _logger.Error(ee); }

                    #region FTP decommissioned
                    // using FTP
                    //if (!filedownloaded)
                    //{
                    //    try
                    //    {
                    //        using (WebClient webClient = new WebClient())
                    //        {
                    //            //Added to login and download from NSE FTP link. 16MAR2021-Amey
                    //            webClient.Credentials = new NetworkCredential(dict_FTPCred["GUEST"].Username, dict_FTPCred["GUEST"].Password);

                    //            webClient.DownloadFile(VaRFileName, VaRExposureFilePath + ExactSpanName);

                    //            filedownloaded = true;
                    //        }
                    //    }
                    //    catch (Exception ee) { _logger.Error(ee); }
                    //}
                    #endregion 
                    // using file
                    try
                    {
                        if (!filedownloaded)
                        {
                            string filename = VaRUrl[2] + "C_VAR1_" + DateTime.Now.STR("ddMMyyyy") + "_" + arr_VaRFileExtensions[VaRIndex] + ".DAT";
                            //Added by Musharraf to Avoid Executing Copy due to invalid Path
                            if (Directory.Exists(VaRExposureFilePath))
                            {
                                File.Copy(filename, VaRExposureFilePath + "C_VAR1_" + DateTime.Now.STR("ddMMyyyy") + "_" + arr_VaRFileExtensions[VaRIndex] + ".DAT", true);
                                filedownloaded = true;
                                //AddToList($"VaR file [{ExactSpanName}] downloaded successfully.");
                            }
                            else
                            {
                                _logger.Debug("File Path Incorrect/Doesn't exist, please change it in Config\nCurrent VaRExposurePath : " + VaRExposureFilePath);
                            }
                        }
                    }
                    catch (Exception ee) { _logger.Error(ee); }

                    _logger.Debug("[VaRExposure] VaRFileName After : " + VaRFileName);

                    if (filedownloaded)
                    {
                        AddToList($"VaR file [{ExactSpanName}] downloaded successfully.");

                        try
                        {
                            arr_VaRFileExtensions = arr_VaRFileExtensions.Take(arr_VaRFileExtensions.Count() - (arr_VaRFileExtensions.Count() - VaRIndex)).ToArray();
                            VaRIndex = 0;
                            _logger.Debug("[VaRExposure] After Slice Index : " + VaRIndex + " arr_VaRFileExtensions Count : " + arr_VaRFileExtensions.Count());
                        }
                        catch (Exception ee) { _logger.Error(ee, "[VaRExposure] Slicing Array"); }
                    }
                    else
                    {
                        VaRIndex++;

                        if (VaRIndex < arr_VaRFileExtensions.Count())
                            DownloadVARExposure(source, e, VaRExposureFilePath);
                        else
                        {
                            VaRIndex = 0;
                            _logger.Debug("[VaRExposure] Failed to dowanload, Index : " + VaRIndex + " arr_VaRFileExtensions Count : " + arr_VaRFileExtensions.Count());
                        }
                    }

                }
            }
            catch (Exception ee)
            {
                VaRIndex = 0;
                _logger.Error(ee);
            }
        }

        private void DownloadBSEExposureFile(string IndexFolderPath, string[] ExposureFolderPath)
        {
            try
            {
                string ExposureFileName = "EF" + DateTime.Now.STR("ddMMyy");
                string ExposureFilePath = ExposureFolderPath[0] + ExposureFileName;/*+ @"\"*/

                AddToList($"BSE-Exposure file [{ExposureFileName}] downloading.");

                using (WebClient webClient = new WebClient())
                    webClient.DownloadFile(ds_Config.GET("URLs", "BSE-EXPOSURE") + ExposureFileName, ExposureFilePath);

                //string NSEExposureFileName = "ael_" + DateTime.Now.STR("ddMMyyyy") + ".csv";
                //string NSEExposureFilePath = ExposureFolderPath + @"\" + NSEExposureFileName;

                var list_ExposureLines = new List<string>();
                using (var stream = new FileStream(ExposureFilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    using (var reader = new StreamReader(stream, Encoding.UTF8))
                    {
                        string line;

                        while ((line = reader.ReadLine()) != null)
                        {
                            // Do something with line, e.g. add to a list or whatever.
                            list_ExposureLines.Add(line);
                        }
                    }
                }

                var dict_BSEUnderlyingNames = ReadBSEConvertorFiles();
                var dict_BSEScripNames = ReadBSEContractFile();

                StringBuilder sb_BSEExposure = new StringBuilder();
                sb_BSEExposure.AppendLine($"#{DateTime.Now.STR("dd-MM-yyyy")}|EF{DateTime.Now.STR("ddMMyy")}|");

                for (int i = 1; i < list_ExposureLines.Count; i++)
                {
                    try
                    {
                        var arr_Fields = list_ExposureLines[i].SPL(',');
                        var Token = Convert.ToInt32(arr_Fields[0]);

                        if (dict_BSEScripNames.ContainsKey(Token) && dict_BSEUnderlyingNames.ContainsKey(dict_BSEScripNames[Token]))
                            sb_BSEExposure.AppendLine($"{Token},{dict_BSEUnderlyingNames[dict_BSEScripNames[Token]]},{arr_Fields[2]}");
                    }
                    catch (Exception ee) { _logger.Error(ee, "DownloadBSEExposure : " + list_ExposureLines[i]); }
                }

                var PrimeDirectory = new DirectoryInfo("C:/Prime");

                string _IndexExposureBSE = "BSEIndexExposure";

                var BSEIndexExposure = PrimeDirectory.GetFiles($"{_IndexExposureBSE}*.csv")
                           .OrderByDescending(f => f.LastWriteTime)
                           .First();

                if (BSEIndexExposure.Length > 0)
                {
                    var arr_IndexExLines = File.ReadAllLines(BSEIndexExposure.FullName);
                    foreach (var line in arr_IndexExLines)
                    {
                        var arr_Fields = line.SPL(',');
                        sb_BSEExposure.AppendLine($"{arr_Fields[0]},{arr_Fields[1]},{arr_Fields[2]}");
                    }
                }

                File.WriteAllText(ExposureFilePath, sb_BSEExposure.STR());

                _logger.Debug("BSE-FOExposure FileName : " + ExposureFileName);

                AddToList($"BSE-Exposure file [{ExposureFileName}] downloaded successfully.");
            }
            catch (Exception expEX)
            {
                _logger.Error(expEX);

                AddToList($"BSE-Exposure file downloaded failed.", true);
            }
        }

        //Added by Akshay on 13-10-2021 for Downloading Exposure File for CD
        private void DownloadCDExposureFile(string IndexExposureFilePath, string[] ExposureFolderPath)
        {
            try
            {
                string ExactFileName = "CDSael_" + DateTime.Now.ToString("ddMMyyyy") + ".csv";
                string ExposureFileName = ExposureFolderPath[0] + ExactFileName;

                AddToList($"CD Exposure file [{ExposureFileName}] downloading.");

                File.WriteAllText(ExposureFileName, File.ReadAllText(IndexExposureFilePath + @"\CDExposure.csv"));

                for (int i = 1; i < ExposureFolderPath.Length; i++)
                    File.Copy(ExposureFolderPath[0] + ExactFileName, ExposureFolderPath[i] + ExactFileName, true);

                AddToList($"CD Exposure file [{ExactFileName}] downloaded successfully.");

                _logger.Debug("CDExposure FileName : " + ExposureFileName);
            }
            catch (Exception ee) { _logger.Error(ee, "DownloadCDSpan"); }
        }

        private Dictionary<string, string> ReadBSEConvertorFiles()
        {
            var dict_BSEUnderlyingNames = new Dictionary<string, string>();

            try
            {
                var PrimeDirectory = new DirectoryInfo("C:/Prime");

                string _ScripConvertorFileFO = "EQD_CC_CO";
                string _ScripConvertorFileCM = "EQ_MAP_CC_";

                var ScripConvertorFileFO = PrimeDirectory.GetFiles($"{_ScripConvertorFileFO}*.csv")
                           .OrderByDescending(f => f.LastWriteTime)
                           .First();

                if (ScripConvertorFileFO.Length > 0)
                {
                    var arr_Lines = File.ReadAllLines(ScripConvertorFileFO.FullName);
                    foreach (var _line in arr_Lines)
                    {
                        try
                        {
                            var arr_Fields = _line.UPP().SPL(',');

                            if (arr_Fields.Length <= 4) continue;

                            if (!dict_BSEUnderlyingNames.ContainsKey(arr_Fields[4].TRM()))
                                dict_BSEUnderlyingNames.Add(arr_Fields[4].TRM(), arr_Fields[3].TRM());
                        }
                        catch (Exception ee) { _logger.Error(ee, "ReadBSEConvertor Loop FO : " + _line); }
                    }
                }

                var ScripConvertorFileCM = PrimeDirectory.GetFiles($"{_ScripConvertorFileCM}*.csv")
                           .OrderByDescending(f => f.LastWriteTime)
                           .First();

                if (ScripConvertorFileCM.Length > 0)
                {
                    var arr_Lines = File.ReadAllLines(ScripConvertorFileCM.FullName);
                    foreach (var _line in arr_Lines)
                    {
                        try
                        {
                            var arr_Fields = _line.UPP().SPL(',');

                            if (arr_Fields.Length <= 5) continue;

                            if (arr_Fields[2].TRM() == "0") continue;

                            if (arr_Fields[6].TRM() != "EQ") continue;

                            if (!dict_BSEUnderlyingNames.ContainsKey(arr_Fields[5].TRM()))
                                dict_BSEUnderlyingNames.Add(arr_Fields[5].TRM(), arr_Fields[2].TRM());
                        }
                        catch (Exception ee) { _logger.Error(ee, "ReadBSEConvertor Loop CM : " + _line); }
                    }
                }


            }
            catch (Exception ee) { _logger.Error(ee); }

            return dict_BSEUnderlyingNames;
        }

        private Dictionary<int, string> ReadBSEContractFile()
        {
            var dict_BSEScripName = new Dictionary<int, string>();

            try
            {
                var PrimeDirectory = new DirectoryInfo("C:/Prime");

                string _ScripFileBSE = "BSE_EQ_SCRIP_";

                var ScripFileBSE = PrimeDirectory.GetFiles($"{_ScripFileBSE}*.txt")
                           .OrderByDescending(f => f.LastWriteTime)
                           .First(); //Modified by Musharraf changed .txt to .csv 26-05-2023

                if (ScripFileBSE.Length > 0)
                {
                    var arr_Lines = File.ReadAllLines(ScripFileBSE.FullName);
                    foreach (var line in arr_Lines)
                    {
                        var arr_Fields = line.SPL('|'); //Modified by Musharraf | to ,
                        if (arr_Fields[1] == "BSE")//Modified by Musharraf index 1 to 54
                        {
                            var Token = Convert.ToInt32(arr_Fields[0]);

                            if (!dict_BSEScripName.ContainsKey(Token))
                                dict_BSEScripName.Add(Token, arr_Fields[2]);
                        }
                    }
                }
            }
            catch (Exception ee) { _logger.Error(ee); }

            return dict_BSEScripName;
        }

        private void DownloadBSESpan(Object source, ElapsedEventArgs e, string[] SpanFilePath)
        {
            try
            {
                _logger.Debug("Outside IF Index : " + BSESpanIndex + " BSESpanFileExtensions Count : " + arr_BSESpanFileExtensions.Count());
                if (BSESpanIndex < arr_BSESpanFileExtensions.Count())
                {
                    //BSERISK20210317-00
                    var SpanFileConfigURL = ds_Config.GET("URLs", "BSE-SPAN");
                    var SpanFileExactName = "BSERISK" + DateTime.Now.Year.STR("0000") + DateTime.Now.Month.STR("00") + DateTime.Now.Date.STR("dd") + "-";
                    SpanFileExactName += arr_BSESpanFileExtensions[BSESpanIndex] + ".spn";

                    var SpanFileName = SpanFileConfigURL + SpanFileExactName;

                    _logger.Debug("SpanFileName Before : " + SpanFileName);

                    using (WebClient webClient = new WebClient())
                        webClient.DownloadFile(SpanFileName, SpanFilePath[0] + SpanFileExactName);

                    _logger.Debug("SpanFileName After : " + SpanFileName);

                    AddToList($"Span file [{SpanFileExactName}] downloaded successfully.");

                    try
                    {
                        arr_BSESpanFileExtensions = arr_BSESpanFileExtensions.Take(arr_BSESpanFileExtensions.Count() - (arr_BSESpanFileExtensions.Count() - BSESpanIndex)).ToArray();
                        BSESpanIndex = 0;
                        _logger.Debug("After Slice Index : " + BSESpanIndex + " BSEFOSpanFileExtensions Count : " + arr_BSESpanFileExtensions.Count());
                    }
                    catch (Exception ee) { _logger.Error(ee, "Slicing Array"); }

                    _logger.Debug("--------------------------------------------------------------------------------");
                }
            }
            catch (WebException)
            {
                BSESpanIndex++;
                _logger.Debug("Inside WebException Before Index : " + BSESpanIndex + " BSEFOSpanFileExtensions Count : " + arr_BSESpanFileExtensions.Count());

                if (BSESpanIndex < arr_BSESpanFileExtensions.Count())
                    DownloadBSESpan(source, e, SpanFilePath);
                else
                {
                    BSESpanIndex = 0;
                    _logger.Debug("Inside WebException After Index : " + BSESpanIndex + " BSEFOSpanFileExtensions Count : " + arr_BSESpanFileExtensions.Count());
                }
            }
            catch (Exception ee)
            {
                BSESpanIndex = 0;
                _logger.Error(ee, "DownloadBSEFOSpan");
            }
        }

        private void DownloadFOSecBanFile(string[] URL, string SavePath)
        {
            _logger.Debug($"Checking license status: EnabledSegments.FO: {_LicenseInfo.EnabledSegments.FO}");
            if (!_LicenseInfo.EnabledSegments.FO)
            {
                return;
            }
            //string downloadFTP = ((string)xmlDoc.Element("BOD-Utility").Element("FO").Element("SECBAN").Element("FTP")).Trim();//To download file FTP
            //string[] arr_SecurityUrl = downloadFTP.Split(',');

            //String URL Parameter is Already passed in the function

            string downloadFromLocal = ((string)xmlDoc.Element("BOD-Utility").Element("FO").Element("SECBAN").Element("LOCAL")).Trim();//To download from local file
            string[] arr_FOSecLocalFile = downloadFromLocal.Split(',');

            var FileName = ((string)xmlDoc.Element("BOD-Utility").Element("FO").Element("SECBAN").Element("NAME")).Trim();
            string[] arr_FileName = FileName.Split(',');
            var dateToCheck = DateTime.Now;
            string FOSecFileName = arr_FileName[0].Replace("$date:ddMMyyyy$", dateToCheck.ToString("ddMMyyyy"));/*$"fo_secban_{DateTime.Now.ToString("ddMMyyyy")}.csv";*/
            try
            {
                bool filedownloaded = false;
                //added on 15APR2021 by Amey. To delete old files.
                var DownloadDirectory = new DirectoryInfo(SavePath);
                try
                {
                    var OldFiles = DownloadDirectory.GetFiles("fo_secban_*.csv");
                    foreach (var OldFOSecfile in OldFiles)
                        File.Delete(OldFOSecfile.FullName);
                }
                catch (Exception ee)
                {
                    _logger.Error(ee);
                }

                AddToList($"FOSecBan file downloading.");

                //if (dateToCheck.DayOfWeek == DayOfWeek.Saturday || dateToCheck.DayOfWeek == DayOfWeek.Sunday)
                //{
                //    // if so, set the date to the previous Friday
                //    dateToCheck = (dateToCheck.DayOfWeek == DayOfWeek.Saturday) ? dateToCheck.AddDays(-1) : dateToCheck.AddDays(-2);
                //}
                for (int i = 0; i < 7; i++)
                {
                    FOSecFileName = arr_FileName[0].Replace("$date:ddMMyyyy$", dateToCheck.ToString("ddMMyyyy"));
                    try
                    {
                        var response = nNSEUtils.Instance.DownloadCommonFile(en_FolderTypes.FO_SEC_BAN, FOSecFileName, SavePath);
                        _logger.Debug("DownloadFOSecBanFile" + FOSecFileName + " | API Response : " + JsonConvert.SerializeObject(response));
                        filedownloaded = response.ResponseStatus == en_ResponseStatus.SUCCESS ? true : false;
                        if (filedownloaded)
                        {
                            //Message is prompt outside the loop
                            FoSecban = FOSecFileName;
                            AddToList($"FO-SEC file [{FOSecFileName}] downloaded successfully.");
                            break;
                        }
                    }
                    catch (Exception ee)
                    {
                        _logger.Error(ee);
                    }

                    //using url
                    try
                    {
                        if (!filedownloaded)
                        {
                            FOSecFileName = arr_FileName[0].Replace("$date:ddMMyyyy$", dateToCheck.ToString("ddMMyyyy"));
                            string month = $"{dateToCheck:MMM}";
                            string url = URL[0] + $"{FOSecFileName}"; ///fo_secban.csv

                            using (WebClient webClient = new WebClient())
                            {
                                webClient.DownloadFile(url, SavePath + FOSecFileName);
                                //using (ZipFile zip = ZipFile.Read(SavePath + @"\" + FOSecFileName))
                                //    zip.ExtractAll(SavePath, ExtractExistingFileAction.DoNotOverwrite);

                                // delete the ZIP file after extraction
                                //File.Delete(Path.Combine(SavePath, FOSecFileName));
                                FoSecban = FOSecFileName;
                            }
                            filedownloaded = true;
                            AddToList($"FO-SEC file [{FOSecFileName}] downloaded successfully.");
                            break;
                        }
                    }
                    catch (Exception ee) { _logger.Error(ee); }
                    #region FTP
                    //if (!filedownloaded)
                    //{

                    //    // using FTP
                    //    try
                    //    {
                    //        using (WebClient webClient = new WebClient())
                    //        {
                    //            //Added to login and download from NSE FTP link. 16MAR2021-Amey
                    //            webClient.Credentials = new NetworkCredential(dict_FTPCred["FO"].Username, dict_FTPCred["FO"].Password);
                    //            webClient.DownloadFile(arr_SecurityUrl[0] + FOSecFileName, SavePath + @"\" + FOSecFileName);
                    //            filedownloaded = true;
                    //        }
                    //    }
                    //    catch (Exception ee) { _logger.Error(ee); }

                    //}
                    #endregion


                    // using file
                    try
                    {
                        if (!filedownloaded)
                        {
                            FOSecFileName = arr_FileName[0].Replace("$date:ddMMyyyy$", dateToCheck.ToString("ddMMyyyy"));
                            string filename = arr_FOSecLocalFile[0] + FOSecFileName;
                            if (File.Exists(filename))
                            {
                                File.Copy(filename, SavePath + FOSecFileName, true);
                                filedownloaded = true;
                                FoSecban = FOSecFileName;
                                AddToList($"FO-SEC file [{FOSecFileName}] downloaded successfully.");
                                break;
                            }
                        }
                    }
                    catch (Exception ee) { _logger.Error(ee); }

                    // subtract a day from the date to check the previous day
                    dateToCheck = dateToCheck.AddDays(-1);

                    // skip weekends
                    if (dateToCheck.DayOfWeek == DayOfWeek.Saturday)
                    {
                        dateToCheck = dateToCheck.AddDays(-1);
                    }
                    else if (dateToCheck.DayOfWeek == DayOfWeek.Sunday)
                    {
                        dateToCheck = dateToCheck.AddDays(-2);
                    }
                }
                if (!filedownloaded)
                    AddToList($"Unable to download FO-SEC file.", true);
                else
                    _logger.Debug($"FO-SEC file [{FOSecFileName}] downloaded successfully.");
            }
            catch (Exception ee) { _logger.Error(ee); AddToList($"Unable to download FO-SEC file.", true); }
        }

        private string DecompressGZAndDelete(FileInfo fileToDecompress, string FileExtension = "")
        {
            string newFileName = "";

            try
            {
                string currentFileName = fileToDecompress.FullName;

                using (FileStream originalFileStream = fileToDecompress.OpenRead())
                {
                    newFileName = currentFileName.Remove(currentFileName.Length - fileToDecompress.Extension.Length) + (FileExtension != "" ? FileExtension : "");

                    File.Delete(newFileName);

                    using (FileStream decompressedFileStream = File.Create(newFileName))
                    {
                        using (GZipStream decompressionStream = new GZipStream(originalFileStream, CompressionMode.Decompress, true))
                        {
                            decompressionStream.CopyTo(decompressedFileStream);
                        }
                    }
                }

                File.Delete(currentFileName);
            }
            catch (Exception ee) { _logger.Error(ee); newFileName = ""; }

            return newFileName;
        }

        private void listBox_Messages_DrawItem(object sender, ListBoxDrawItemEventArgs e)
        {
            try
            {
                if (hs_ErrorIndex.CON(listBox_Messages.Items[e.Index].STR()))
                    e.Appearance.ForeColor = Color.OrangeRed;
            }
            catch (Exception ee) { _logger.Error(ee); }
        }

        // Added by Snehadri on 15JUN2021 for Automatic BOD Process
        private void ReadConfig()
        {
            try
            {
                XmlTextReader tReader = new XmlTextReader("C:/Prime/Config.xml");
                tReader.Read();
                ds_SQLConfig.ReadXml(tReader);

                var DBInfo = ds_SQLConfig.Tables["DB"].Rows[0];

                //added convert zero datetime=True on 16APR2021 by Amey. Was not assigning SQL datetime to C# DateTime in sp_ContractMaster.
                _MySQLCon = $"Data Source={DBInfo["SERVER"]};Initial Catalog={DBInfo["NAME"]};UserID={DBInfo["USER"]};Password={DBInfo["PASSWORD"]};SslMode=none;convert zero datetime=True;";

                var CONInfo = ds_SQLConfig.Tables["CONNECTION"].Rows[0];

                SetMaxAllowedSqlPacket();

                //added on 28FEB2021 by Amey
                FlushMySQLConnectionErrors();

                //UseUdiffFormat = Convert.ToBoolean(ds_Config.Tables["OTHER"].Rows[0]["USE-UDIFF-FORMAT"].ToString());
                isMcxContractFileContainsHeader = Convert.ToBoolean(ds_Config.Tables["OTHER"].Rows[0]["MCX-CONTRACT-CONTAINS-HEADERS"].ToString());


            }
            catch (Exception error)
            {
                _logger.Error(error, ": ReadConfig");
                XtraMessageBox.Show("Invalid entry in Config file. Please check logs for more details.", "Error");
            }
        }

        // Added by Snehadri on 15JUN2021 for Automatic BOD Process
        private void SetMaxAllowedSqlPacket()
        {
            try
            {
                using (MySqlConnection myConn = new MySqlConnection(_MySQLCon))
                {
                    //changed to SP on 27APR2021 by Amey
                    using (MySqlCommand myCmd = new MySqlCommand("sp_SetMaxAllowedPacket", myConn))
                    {
                        myCmd.CommandType = CommandType.StoredProcedure;

                        myConn.Open();
                        myCmd.ExecuteNonQuery();
                        myConn.Close();
                    }
                }
            }
            catch (Exception errror)
            {
                _logger.Error(errror, "SetMaxAllowedSqlPacket");
            }
        }

        // Added by Snehadri on 15JUN2021 for Automatic BOD Process
        private void FlushMySQLConnectionErrors()
        {
            try
            {
                using (MySqlConnection myConn = new MySqlConnection(_MySQLCon))
                {
                    using (MySqlCommand myCmd = new MySqlCommand("flush hosts;", myConn))
                    {
                        myConn.Open();
                        myCmd.ExecuteNonQuery();
                        myConn.Close();
                    }
                }
            }
            catch (Exception errror)
            {
                _logger.Error(errror, " : FlushMySQLConnectionErrors");
            }
        }

        // Added by Snehadri on 14JUL2021
        private void CheckMaxAllowedSqlPacket()
        {
            string _pcket = string.Empty;
            try
            {
                using (MySqlConnection myCon = new MySqlConnection(_MySQLCon))
                {
                    //changed to SP on 27APR2021 by Amey
                    using (MySqlCommand cmd = new MySqlCommand("sp_GetMaxAllowedPacket", myCon))
                    {
                        myCon.Open();

                        cmd.CommandType = CommandType.StoredProcedure;

                        using (MySqlDataReader reader = cmd.ExecuteReader())
                        {
                            while (reader.Read())
                                _pcket = reader.GetString(1);

                            reader.Close();
                        }
                        myCon.Close();
                    }
                }
                if (_pcket != "1073741824")
                {
                    Application.Restart();
                    Environment.Exit(0);
                }
            }
            catch (Exception errror)
            {
                _logger.Error(errror, "CheckMaxAllowedSqlPacket ");
            }
        }

        // Added by Snehadri on 15JUN2021 for Automatic BOD Process
        private void StartCMFeedReceivers(string[] CMFeedPath)
        {
            // CM Feed Receiver     
            try
            {
                CloseComponentexe("CM FeedReceiver");

                foreach (var cmpath in CMFeedPath)
                {
                    OpenComponentexe("CM FeedReceiver", cmpath);
                    bool connected = GatewayEngineConnector.ConnectComponents("CMFeedReceiver");

                    if (connected)
                    {
                        AddToList($"CM Feed Receiver Started ({cmpath}).");
                        if (!list_ComponentStarted.Contains("CMFeedReceiver")) { list_ComponentStarted.Add("CMFeedReceiver"); }
                        Thread.Sleep(3000);
                    }
                    else
                    {
                        CloseComponentexe("CM FeedReceiver");

                        AddToList($"Can't open CM feed receiver ({cmpath}). Please check log.", true);
                        IsWorking = false;
                        btn_RestartAuto.Enabled = true;
                        btn_Settings.Enabled = true;
                        SentMail("CM Feed Receiver has failed to launch.");

                        if (list_ComponentStarted.Contains("CMFeedReceiver")) { list_ComponentStarted.Remove("CMFeedReceiver"); }
                    }
                }
            }
            catch (Exception ee)
            {
                _logger.Error(ee, "CM Feed Receiver");
                AddToList("Can't open CM feed receiver. Please check log", true);
                IsWorking = false;
                btn_RestartAuto.Enabled = true;
                btn_Settings.Enabled = true;
                if (list_ComponentStarted.Contains("CMFeedReceiver")) { list_ComponentStarted.Remove("CMFeedReceiver"); }

                SentMail("CM Feed Receiver has failed to launch.");

            }
        }

        private void StartFOFeedReceiver(string[] FOFeedpath)
        {

            try
            {
                CloseComponentexe("FO FeedReceiver");

                foreach (var fopath in FOFeedpath)
                {
                    OpenComponentexe("FO FeedReceiver", fopath);
                    bool connected = GatewayEngineConnector.ConnectComponents("FOFeedReceiver");
                    if (connected)
                    {
                        AddToList($"FO Feed Receiver Started ({fopath}).");
                        if (!list_ComponentStarted.Contains("FOFeedReceiver"))
                            list_ComponentStarted.Add("FOFeedReceiver");
                        Thread.Sleep(3000);
                    }
                    else
                    {
                        CloseComponentexe("FO FeedReceiver");
                        AddToList("Can't open FO feed receiver. Please check log", true);
                        IsWorking = false;
                        btn_RestartAuto.Enabled = true;
                        btn_Settings.Enabled = true;
                        SentMail("FO Feed Receiver has failed to launch");
                        if (list_ComponentStarted.Contains("FOFeedReceiver"))
                            list_ComponentStarted.Remove("FOFeedReceiver");
                        break;
                    }
                }
            }
            catch (Exception ee)
            {
                _logger.Error(ee, "FO Feed Receiver");
                AddToList("Can't open FO feed receiver. Please check log", true);
                IsWorking = false;
                btn_RestartAuto.Enabled = true;
                btn_Settings.Enabled = true;
                if (list_ComponentStarted.Contains("FOFeedReceiver"))
                    list_ComponentStarted.Remove("FOFeedReceiver");
                SentMail("FO Feed Receiver has failed to launch");
            }
        }

        private void StartCDFeedReceiver(string[] CDFeedPath)
        {
            // CD Feed Receiver   
            try
            {
                CloseComponentexe("CD FeedReceiver");

                foreach (var cdpath in CDFeedPath)
                {
                    OpenComponentexe("CD FeedReceiver", cdpath);

                    bool connected = GatewayEngineConnector.ConnectComponents("CDFeedReceiver");

                    if (connected)
                    {
                        AddToList($"CD Feed Receiver Started ({cdpath}).");
                        if (!list_ComponentStarted.Contains("CDFeedReceiver")) { list_ComponentStarted.Add("CDFeedReceiver"); }
                        Thread.Sleep(3000);
                    }
                    else
                    {
                        CloseComponentexe("CD FeedReceiver");
                        AddToList("Can't open CD feed receiver. Please check log.", true);
                        IsWorking = false;
                        btn_RestartAuto.Enabled = true;
                        SentMail("CD Feed Receiver has failed to launch.");
                        if (list_ComponentStarted.Contains("CDFeedReceiver"))
                            list_ComponentStarted.Remove("CDFeedReceiver");

                    }
                }
            }
            catch (Exception ee)
            {
                _logger.Error(ee, "CD Feed Receiver");
                AddToList("Can't open CD feed receiver. Please check log", true);
                IsWorking = false;
                btn_RestartAuto.Enabled = true;
                if (list_ComponentStarted.Contains("CDFeedReceiver")) { list_ComponentStarted.Remove("CDFeedReceiver"); }
                SentMail("CD Feed Receiver has failed to launch.");
            }

        }

        // Added by Snehadri on 15JUN2021 for Automatic BOD Process
        private void StartCMNotisApi(string[] NOTISEQPath)
        {
            // NOTIS EQ Receiver
            try
            {
                CloseComponentexe("NOTIS API EQ Manager");

                foreach (var path in NOTISEQPath)
                {
                    OpenComponentexe("NOTIS EQ Receiver", path);
                    bool notiscdstarted = GatewayEngineConnector.ConnectComponents("NOTISEQ");
                    if (notiscdstarted)
                    {
                        AddToList("NOTIS EQ Started");
                        Thread.Sleep(5000);
                        if (!list_ComponentStarted.Contains("NOTISEQReceiver"))
                            list_ComponentStarted.Add("NOTISEQReceiver");
                    }
                    else
                    {
                        CloseComponentexe("NOTIS API EQ Manager");
                        AddToList("NOTIS EQ has failed to launch. Please check the log", true);
                        IsWorking = false;
                        btn_RestartAuto.Enabled = true;
                        btn_Settings.Enabled = true;
                        if (list_ComponentStarted.Contains("NOTISEQReceiver"))
                            list_ComponentStarted.Remove("NOTISEQReceiver");
                        SentMail("NOTIS EQ has failed to launch.");

                    }
                }
            }
            catch (Exception error)
            {
                AddToList("NOTIS EQ has failed to launch. Please check the log", true);
                _logger.Error(error);
                IsWorking = false;
                btn_RestartAuto.Enabled = true;
                btn_Settings.Enabled = true;
                if (list_ComponentStarted.Contains("NOTISEQReceiver"))
                    list_ComponentStarted.Remove("NOTISEQReceiver");
                SentMail("NOTIS EQ has failed to launch.");
            }

        }

        private void StartFONotisApi(string[] NOTISFOPath)
        {
            // NOTIS FO Receiver
            try
            {
                CloseComponentexe("NOTIS API FO Manager");

                foreach (var path in NOTISFOPath)
                {
                    OpenComponentexe("NOTIS FO Receiver", path);
                    bool notiscdstarted = GatewayEngineConnector.ConnectComponents("NOTISFO");
                    if (notiscdstarted)
                    {
                        AddToList("NOTIS FO Started");
                        Thread.Sleep(5000);
                        if (!list_ComponentStarted.Contains("NOTISFOReceiver"))
                            list_ComponentStarted.Add("NOTISFOReceiver");
                    }
                    else
                    {
                        CloseComponentexe("NOTIS API FO Manager");
                        AddToList("NOTIS FO has failed to launch. Please check the log", true);
                        IsWorking = false;
                        btn_RestartAuto.Enabled = true;
                        btn_Settings.Enabled = true;
                        if (list_ComponentStarted.Contains("NOTISFOReceiver"))
                            list_ComponentStarted.Remove("NOTISFOReceiver");
                        SentMail("NOTIS FO has failed to launch.");
                    }
                }
            }
            catch (Exception error)
            {
                _logger.Error(error);
                AddToList("NOTIS FO has failed to launch. Please check the log", true);
                IsWorking = false;
                btn_RestartAuto.Enabled = true;
                btn_Settings.Enabled = true;
                if (list_ComponentStarted.Contains("NOTISFOReceiver"))
                    list_ComponentStarted.Remove("NOTISFOReceiver");
                SentMail("NOTIS FO has failed to launch.");
            }
        }

        private void StartCDNotisApi(string[] NOTISCDPath)
        {

            try
            {
                CloseComponentexe("NOTIS API CD Manager");

                foreach (var path in NOTISCDPath)
                {
                    OpenComponentexe("NOTIS CD Receiver", path);
                    bool notiscdstarted = GatewayEngineConnector.ConnectComponents("NOTISCD");
                    if (notiscdstarted)
                    {
                        AddToList("NOTIS CD Started");
                        Thread.Sleep(5000);
                        if (!list_ComponentStarted.Contains("NOTISCDReceiver"))
                            list_ComponentStarted.Add("NOTISCDReceiver");
                    }
                    else
                    {
                        CloseComponentexe("NOTIS API CD Manager");
                        AddToList("NOTIS CD has failed to launch. Please check the log", true);
                        IsWorking = false;
                        btn_RestartAuto.Enabled = true;
                        btn_Settings.Enabled = true;
                        if (list_ComponentStarted.Contains("NOTISCDReceiver"))
                            list_ComponentStarted.Remove("NOTISCDReceiver");
                        SentMail("NOTIS CD has failed to launch.");
                    }
                }

            }
            catch (Exception error)
            {
                _logger.Error(error);
                AddToList("NOTIS CD has failed to launch. Please check the log", true);
                IsWorking = false;
                btn_RestartAuto.Enabled = true;
                btn_Settings.Enabled = true;
                if (list_ComponentStarted.Contains("NOTISCDReceiver"))
                    list_ComponentStarted.Remove("NOTISCDReceiver");
                SentMail("NOTIS CD has failed to launch.");
            }
        }

        // Added by Snehadri on 15JUN2021 for Automatic BOD Process
        private void InsertTokensIntoDB()
        {
            AddToList("Upload Token Started");
            bool CMTokenUploaded = false; bool FOTokenUploaded = false;

            try
            {
                List<string> list_ContractMasterRows = new List<string>();
                List<Exchange.Security> list_EQSecurity = new List<Exchange.Security>();

                StringBuilder sb_InsertCommand = new StringBuilder("");

                #region IndexTokens and IndexScrip Upload
                //IndexTokens.csv
                try
                {
                    sb_InsertCommand.Append("TRUNCATE tbl_indextokens; ");
                    // Construct file path for the CSV file
                    string IndexTokens = "IndexTokens.csv";
                    string csvFilePath = @"C:\Prime\" + IndexTokens;

                    // Check if the CSV file exists
                    if (File.Exists(csvFilePath))
                    {
                        // StringBuilder to store the INSERT command
                        sb_InsertCommand.Append("INSERT IGNORE INTO tbl_indextokens (Symbol, Token, FullName, Segment) VALUES ");

                        // Read the CSV file into a list of strings
                        List<string> csvLines = File.ReadAllLines(csvFilePath).ToList();

                        // Loop through each line in the CSV file
                        foreach (string line in csvLines)
                        {
                            //if (line.Contains("Security"))
                            //{ continue; }
                            // Split the line into its individual values
                            string[] values = line.Split(',');

                            // Construct the INSERT command and add it to the StringBuilder
                            sb_InsertCommand.Append($"('{values[0]}', {values[1]}, '{values[2]}', '{values[3]}'),");
                        }

                        // Remove the last comma from the INSERT command
                        sb_InsertCommand.Remove(sb_InsertCommand.Length - 1, 1);

                        // Create a new MySqlConnection object and open the connection
                        using (MySqlConnection conn = new MySqlConnection(_MySQLCon))
                        {
                            conn.Open();

                            // Create a new MySqlCommand object and set its properties
                            MySqlCommand cmd = new MySqlCommand(sb_InsertCommand.ToString(), conn);
                            cmd.CommandType = CommandType.Text;

                            // Execute the INSERT command
                            cmd.ExecuteNonQuery();

                            // Close the connection
                            conn.Close();
                        }
                    }
                    else
                    {
                        AddToList($"IndexToken file Missing/Incorrect Format {IndexTokens} in {csvFilePath},couldn't be uploaded in DB", true);
                    }
                }
                catch (Exception ex)
                {
                    _logger.Error(ex, " : InsertTokensIntoDB IndexToken not Uploaded to DB check tbl_indextokens");
                }

                sb_InsertCommand.Clear();

                //IndexScrip.csv
                try
                {
                    sb_InsertCommand.Append("TRUNCATE tbl_indexscrip; ");
                    // Construct file path for the CSV file
                    string IndexScrip = "IndexScrip.csv";
                    string csvFilePath = @"C:\Prime\" + IndexScrip;

                    // Check if the CSV file exists
                    if (File.Exists(csvFilePath))
                    {
                        // StringBuilder to store the INSERT command
                        sb_InsertCommand.Append("INSERT IGNORE INTO tbl_indexscrip (FullName, Symbol) VALUES ");

                        // Read the CSV file into a list of strings
                        List<string> csvLines = File.ReadAllLines(csvFilePath).ToList();

                        // Loop through each line in the CSV file
                        foreach (string line in csvLines)
                        {
                            //if (line.Contains("Security"))
                            //{ continue; }
                            // Split the line into its individual values
                            string[] values = line.Split(',');

                            // Construct the INSERT command and add it to the StringBuilder
                            sb_InsertCommand.Append($"('{values[0]}', '{values[1]}'),");
                        }

                        // Remove the last comma from the INSERT command
                        sb_InsertCommand.Remove(sb_InsertCommand.Length - 1, 1);

                        // Create a new MySqlConnection object and open the connection
                        using (MySqlConnection conn = new MySqlConnection(_MySQLCon))
                        {
                            conn.Open();

                            // Create a new MySqlCommand object and set its properties
                            MySqlCommand cmd = new MySqlCommand(sb_InsertCommand.ToString(), conn);
                            cmd.CommandType = CommandType.Text;

                            // Execute the INSERT command
                            cmd.ExecuteNonQuery();

                            // Close the connection
                            conn.Close();
                        }
                    }
                    else
                    {
                        AddToList($"IndexScrip file Missing/Incorrect Format {IndexScrip} in {csvFilePath},couldn't be uploaded in DB", true);
                    }
                }
                catch (Exception ex)
                {
                    _logger.Error(ex, " : InsertTokensIntoDB IndexScrip not Uploaded to DB check tbl_indexscrip");
                }

                sb_InsertCommand.Clear();

                #endregion
                sb_InsertCommand.Append("TRUNCATE tbl_contractmaster; ");
                #region CM Security

                try
                {
                    //CM_security_fileName = "NSE_CM_security_22062023.csv";
                    string CMBhavCopyFile = _NSE_CM_bhavcopy;
                    var BhavcopyFile = new DirectoryInfo(@"C:\Prime\").GetFiles("NSE_CM_Bhavcopy*.csv").OrderByDescending(x => x.LastWriteTime).FirstOrDefault();// @"C:\Prime\" + CMBhavCopyFile;
                    var Bhavcopy = new List<CMBhavcopy>();

                    if (BhavcopyFile != null)
                    {
                      Bhavcopy = Exchange.ReadCMBhavcopy(BhavcopyFile.FullName, true);//To fetch Closing and Settlement Price added by Musharra0f
                    }
                    
                    //added Exists check on 27APR2021 by Amey       //security.txt
                    if (File.Exists("C:\\Prime\\" + CM_security_fileName))
                    {
                        sb_InsertCommand.Append("INSERT IGNORE INTO tbl_contractmaster(Token,Symbol,InstrumentName,Series,Segment,ScripName,CustomScripName,ScripType,ExpiryUnix,StrikePrice,LotSize,UnderlyingToken,UnderlyingSegment,ClosingPrice,SettlementPrice,Isin) VALUES");//,ClosingPrice,SettlementPrice added by Musharraf

                        var list_Security = Exchange.ReadSecurity("C:\\Prime\\" + CM_security_fileName, true);//security.txt
                        list_EQSecurity = list_Security.Where(v => v.Series == "EQ").ToList();

                        foreach (var _Security in list_Security)
                        {
                            //var bhavcopyDetails = Bhavcopy.FirstOrDefault(b => b.ScripName == _Security.ScripName && b.Series == _Security.Series);

                            var bhavcopyDetails = Bhavcopy.Where(v => v.CustomScripname == _Security.CustomScripname).FirstOrDefault();//To fetch as per scripname
                            var ClosingPrice = (bhavcopyDetails != null) ? bhavcopyDetails.Close : 0;

                            list_ContractMasterRows.Add($"({_Security.Token},'{_Security.Symbol}','EQ','{_Security.Series}','NSECM','{_Security.ScripName}'," +
                                $"'{_Security.CustomScripname}','EQ','{_Security.ExpiryUnix}',{0},{_Security.LotSize},{_Security.Token},'NSECM',{ClosingPrice},{0},'{_Security.ISIN}')");//bhavcopyDetails.Close

                            if (ClosingPrice <= 0)
                            {
                                _logger.Debug($"closing skipped for : {_Security.CustomScripname}");
                            }
                           
                        }

                        ConcurrentDictionary<string, double> dict_IndexClosing = new ConcurrentDictionary<string, double>();
                        ConcurrentDictionary<string, string> dict_IndexScrips = new ConcurrentDictionary<string, string>();
                        ConcurrentDictionary<string, double> dict_bseIndexClosing = new ConcurrentDictionary<string, double>();
                        string SnapShotFile = DailySnapshot;
                        string IndCloseFilePath = @"C:\Prime\" + SnapShotFile;
                        
                        try
                        {
                            var FilePath = "C://Prime//IndexScrip.csv";
                            if (File.Exists(FilePath))
                            {
                                string[] arr_Lines = File.ReadAllLines(FilePath);
                                foreach (var Line in arr_Lines)
                                {
                                    try
                                    {
                                        string[] arr_Fields = Line.Split(',');
                                        if (!dict_IndexScrips.ContainsKey(arr_Fields[0].Trim().ToUpper()))
                                            dict_IndexScrips.TryAdd(arr_Fields[0].Trim().ToUpper(), arr_Fields[1].Trim().ToUpper());
                                    }
                                    catch(Exception ee) { _logger.Error(ee); }                                   
                                }
                            }
                        }
                        catch (Exception ee) { _logger.Error(ee); }

                        try
                        {
                            if (File.Exists(IndCloseFilePath))
                            {
                                foreach (var item in File.ReadAllLines(IndCloseFilePath))
                                {
                                    try
                                    {
                                        var arr_fields = item.Split(',');
                                        var IndexName = arr_fields[0].Trim().ToUpper();
                                        var closingPrice = Convert.ToDouble(arr_fields[5]);
                                        if (dict_IndexScrips.TryGetValue(IndexName, out string Symbol))
                                            dict_IndexClosing.TryAdd(Symbol, closingPrice);
                                    }
                                    catch (Exception ee) { _logger.Error(ee); }
                                }
                            }
                        }
                        catch(Exception ee) { _logger.Error(ee); }
                        
                        try
                        {
                            var BseIndexCloseFile = new DirectoryInfo(@"C:\Prime\").GetFiles("index5_*").OrderByDescending(x => x.LastWriteTime).FirstOrDefault();
                            if (BseIndexCloseFile != null)
                            {
                                foreach (var line in File.ReadAllLines(BseIndexCloseFile.FullName))
                                {
                                    try
                                    {
                                        var arr_Fields = line.Split('|');
                                        var closing = Convert.ToDouble(arr_Fields[1]) / 100;
                                        dict_bseIndexClosing.TryAdd(arr_Fields[0], closing);
                                    }
                                    catch (Exception ee) { _logger.Error(ee); }

                                }
                            }
                        }
                        catch(Exception ee) { _logger.Error(ee); }
                       
                        if (File.Exists("C:\\Prime\\IndexTokens.csv"))
                        {
                            foreach (var item in File.ReadAllLines("C:\\Prime\\IndexTokens.csv"))
                            {
                               
                                string[] arr_Fields = item.Split(',');
                                string IndexName = arr_Fields[0];
                                string CustomScripName = $"{IndexName}|0|EQ|0";
                                double closingPrice = 0;
                                if (arr_Fields[3] == "BSE")
                                    dict_bseIndexClosing.TryGetValue(arr_Fields[4],out closingPrice);
                                else
                                    dict_IndexClosing.TryGetValue(IndexName, out closingPrice);
                                
                                list_ContractMasterRows.Add($"({arr_Fields[1]},'{IndexName}','EQ','{en_InstrumentName.EQ}','{arr_Fields[3]}CM','{IndexName}-EQ'," +
                                      $"'{CustomScripName}','EQ',{0},{0},{1},{arr_Fields[1]},'{arr_Fields[3]}CM',{closingPrice},{0},'')");

                            }
                        }
                        else
                            XtraMessageBox.Show("Index Token file is not available.", "Error");

                        sb_InsertCommand.Append(string.Join(",", list_ContractMasterRows));

                        using (MySqlConnection myConnToken = new MySqlConnection(_MySQLCon))
                        {
                            using (MySqlCommand myCmd = new MySqlCommand(sb_InsertCommand.ToString(), myConnToken))
                            {
                                myConnToken.Open();
                                myCmd.CommandType = CommandType.Text;
                                myCmd.ExecuteNonQuery();
                                myConnToken.Close();
                            }
                        }
                        list_ComponentStarted.Add("CMUploadToken");
                        CMTokenUploaded = true;
                    }
                    else
                    {
                        AddToList("Security file not found", true);
                        CMTokenUploaded = false;
                    }
                }
                catch (Exception ee)
                {
                    _logger.Error(ee, " : InsertTokensIntoDB CM");

                    AddToList("CM Tokens Upload failed. Please check logs for more details.", true);

                    IsWorking = false;
                    btn_RestartAuto.Enabled = true;
                    btn_Settings.Enabled = true;
                    CMTokenUploaded = false;

                    if (list_ComponentStarted.Contains("CMUploadToken"))
                        list_ComponentStarted.Remove("CMUploadToken");
                    if (list_ComponentStarted.Contains("FOUploadToken"))
                        list_ComponentStarted.Remove("FOUploadToken");

                    SentMail("CM Tokens");

                }

                #endregion

                #region FO Contract

                list_ContractMasterRows.Clear();
                sb_InsertCommand.Clear();

                try
                {

                    string FOBhavCopyFile = _FOBhavcopy.EndsWith(".zip") ? _FOBhavcopy.Substring(0, _FOBhavcopy.Length - 4) : _FOBhavcopy.EndsWith(".csv") ? _FOBhavcopy : throw new ArgumentException("Invalid file extension");
                    string FilePath = FOBhavCopyFile.Contains("C:\\Prime\\") ? FOBhavCopyFile : (@"C:\Prime\" + FOBhavCopyFile);

                    var Bhavcopy = Exchange.ReadFOBhavcopy(FilePath, true);//To fetch Closing and Settlement Price added by Musharraf

                    //added Exists check on 27APR2021 by Amey   //contract.txt
                    if (File.Exists("C:\\Prime\\" + FO_contract_fileName))
                    {
                        sb_InsertCommand = new StringBuilder("INSERT IGNORE INTO tbl_contractmaster(Token,Symbol,InstrumentName,Series,Segment,ScripName,CustomScripName,ScripType,ExpiryUnix,StrikePrice,LotSize,UnderlyingToken,UnderlyingSegment,ClosingPrice,SettlementPrice,OpenInterest) VALUES");//Closing and Settlement added by Musharraf

                        var list_Contract = Exchange.ReadContract("C:\\Prime\\" + FO_contract_fileName, true); //contract.txt
                        var FutContracts = list_Contract.Where(v => v.ScripType == en_ScripType.XX).ToList();

                        foreach (var _Contract in list_Contract)
                        {

                            var USegment = "NSEFO";
                            var UnderlyingToken = -1;

                            double ClosingPrice =  0;
                            double SettlingPrice = 0;
                            long OpenInterest =  0;

                            if (_Contract.ScripType == en_ScripType.XX)
                                UnderlyingToken = _Contract.Token;
                            else
                            {
                                var temp = FutContracts.Where(v => v.Symbol == _Contract.Symbol && v.Expiry.Month == _Contract.Expiry.Month && v.Expiry.Year == _Contract.Expiry.Year).FirstOrDefault();
                                if (temp != null)
                                    UnderlyingToken = temp.Token;
                                else
                                {
                                    //var twmp = list_EQSecurity.Where(v => v.Symbol == _Contract.Symbol).FirstOrDefault();
                                    //if (twmp != null)
                                    //{
                                    //    UnderlyingToken = twmp.Token;
                                    //    USegment = "NSECM";
                                    //}

                                    var temp1 = FutContracts.Where(v => v.Symbol == _Contract.Symbol).OrderByDescending(v => v.Expiry).FirstOrDefault();

                                    if (temp1 != null)
                                        UnderlyingToken = temp1.Token;

                                }
                            }

                            var pricefrombhavcopy = Bhavcopy.Where(v => v.CustomScripname == _Contract.CustomScripname).FirstOrDefault();//Added by Musharraf to fetch closing and settling price
                            if (pricefrombhavcopy != null)
                            {
                                 ClosingPrice =  pricefrombhavcopy.Close; //if null value is entered then Default should be zero
                                 SettlingPrice =  pricefrombhavcopy.SettlePrice;
                                 OpenInterest = pricefrombhavcopy.OpenInterest;
                            }
                            //var ClosingPrice = (pricefrombhavcopy != null) ? pricefrombhavcopy.Close : 0; //if null value is entered then Default should be zero
                            //var SettlingPrice = (pricefrombhavcopy != null) ? pricefrombhavcopy.SettlePrice : 0;
                            //var OpenInterest = (pricefrombhavcopy != null) ? pricefrombhavcopy.OpenInterest : 0;
                            if (UnderlyingToken != -1)
                                list_ContractMasterRows.Add($"({_Contract.Token},'{_Contract.Symbol}','{_Contract.Instrument}','-','NSEFO','{_Contract.ScripName}'," +
                                    $"'{_Contract.CustomScripname}','{_Contract.ScripType}','{_Contract.ExpiryUnix}',{_Contract.StrikePrice},{_Contract.LotSize}," +
                                    $"{UnderlyingToken},'{USegment}',{ClosingPrice},{SettlingPrice},{OpenInterest})");//pricefrombhavcopy.Close, pricefrombhavcopy.SettlePrice
                        }

                        sb_InsertCommand.Append(string.Join(",", list_ContractMasterRows));

                        using (MySqlConnection myConnToken = new MySqlConnection(_MySQLCon))
                        {
                            using (MySqlCommand myCmd = new MySqlCommand(sb_InsertCommand.ToString(), myConnToken))
                            {
                                myConnToken.Open();
                                myCmd.CommandType = CommandType.Text;
                                myCmd.ExecuteNonQuery();
                                myConnToken.Close();
                            }
                        }

                        list_ComponentStarted.Add("FOUploadToken");
                        FOTokenUploaded = true;
                    }
                    else
                    {
                        AddToList("Contract file not found", true);
                        FOTokenUploaded = false;
                    }
                }
                catch (Exception ee)
                {
                    _logger.Error(ee, " : InsertTokensIntoDB FO");

                    AddToList("FO Tokens Upload failed. Please check logs for more details.", true);

                    IsWorking = false;
                    btn_RestartAuto.Enabled = true;
                    btn_Settings.Enabled = true;
                    FOTokenUploaded = false;

                    if (list_ComponentStarted.Contains("CMUploadToken"))
                        list_ComponentStarted.Remove("CMUploadToken");
                    if (list_ComponentStarted.Contains("FOUploadToken"))
                        list_ComponentStarted.Remove("FOUploadToken");


                    SentMail("FO Tokens");
                }

                #endregion

                #region CD Contract

                list_ContractMasterRows.Clear();
                sb_InsertCommand.Clear();

                try
                {

                    string path = "C:\\Prime\\" + _CDBhavcopy;//Added by Musharraf 
                    var CdBhavCopyLines = Exchange.ReadCDBhavcopy(path, true);//to Fetch Closing Price
                    //added Exists check on 27APR2021 by Amey    //cd_contract.txt
                    if (File.Exists("C:\\Prime\\" + CD_contract_fileName))
                    {
                        sb_InsertCommand = new StringBuilder("INSERT IGNORE INTO tbl_contractmaster(Token,Symbol,InstrumentName,Series,Segment,ScripName,CustomScripName,ScripType,ExpiryUnix,StrikePrice,LotSize,UnderlyingToken,UnderlyingSegment,ClosingPrice,SettlementPrice,Multiplier,PrevClosingPrice,CD_OpenInterest) VALUES");//Closing and Settlement added by Musharraf

                        var list_CDContract = Exchange.ReadCDContract("C:\\Prime\\" + CD_contract_fileName, true);//cd_contract.txt
                        var FutContracts = list_CDContract.Where(v => v.ScripType == en_ScripType.XX).ToList();

                        foreach (var _CDContract in list_CDContract)
                        {
                            var UnderlyingToken = -1;
                            if (_CDContract.ScripType == en_ScripType.XX)
                                UnderlyingToken = _CDContract.Token;
                            else
                            {
                                var temp = FutContracts.Where(v => v.Symbol == _CDContract.Symbol && v.Expiry.Month == _CDContract.Expiry.Month && v.Expiry.Year == _CDContract.Expiry.Year).FirstOrDefault();
                                if (temp != null)
                                    UnderlyingToken = temp.Token;
                            }

                            var cdbhavcopydetails = CdBhavCopyLines.FirstOrDefault(v => v.Symbol == _CDContract.Symbol && v.Expiry.Date == _CDContract.Expiry.Date && v.StrikePrice == _CDContract.StrikePrice && v.ScripType == _CDContract.ScripType);//To fetch closing Price
                            var ClosingPrice = (cdbhavcopydetails == null) ? 0 : cdbhavcopydetails.Close; //Add a null check for closing price
                            var PreviousClosPrice = (cdbhavcopydetails == null) ? 0 : cdbhavcopydetails.PreviousClose;
                            var CD_OpenInterest = (cdbhavcopydetails == null) ? 0 : cdbhavcopydetails.OpenInterest;

                            if (UnderlyingToken != -1)
                                list_ContractMasterRows.Add($"({_CDContract.Token},'{_CDContract.Symbol}','{_CDContract.Instrument}','-','NSECD','{_CDContract.ScripName}'," +
                                    $"'{_CDContract.CustomScripname}','{_CDContract.ScripType}','{_CDContract.ExpiryUnix}',{_CDContract.StrikePrice},{_CDContract.LotSize}," +
                                    $"{UnderlyingToken},'NSECD',{ClosingPrice},{0},{_CDContract.Multiplier},{PreviousClosPrice},{CD_OpenInterest})");
                        }

                        sb_InsertCommand.Append(string.Join(",", list_ContractMasterRows));

                        using (MySqlConnection myConnToken = new MySqlConnection(_MySQLCon))
                        {
                            using (MySqlCommand myCmd = new MySqlCommand(sb_InsertCommand.ToString(), myConnToken))
                            {
                                myConnToken.Open();
                                myCmd.CommandType = CommandType.Text;
                                myCmd.ExecuteNonQuery();
                                myConnToken.Close();
                            }
                        }
                    }
                }
                catch (Exception ee)
                {
                    _logger.Error(ee, " : InsertTokensIntoDB CD");

                    AddToList("CD Tokens Upload failed. Please check logs for more details.", true);
                }

                #endregion

                #region BSECM Security

                try
                {
                    list_ContractMasterRows.Clear();
                    sb_InsertCommand.Clear();
                    sb_InsertCommand.Append("INSERT IGNORE INTO tbl_contractmaster(Token,Symbol,InstrumentName,Series,Segment,ScripName,CustomScripName,ScripType,ExpiryUnix,StrikePrice,LotSize,UnderlyingToken,UnderlyingSegment,ClosingPrice,PrevClosingPrice,Isin,TotalTrades) VALUES");

                    DirectoryInfo _PrimeDirectory = new DirectoryInfo("C:\\Prime\\");
                    var BSESecurity = _PrimeDirectory.GetFiles("BSE_EQ_SCRIP_*.csv").OrderByDescending(v => v.LastWriteTime).ToList();//changed .txt to .csv by Musharraf 21st April 2023

                    var BsecmBhavPath = _PrimeDirectory.GetFiles("BSE_EQ_BHAVCOPY_*.csv").OrderByDescending(v => v.LastWriteTime).ToList();

                    if (BSESecurity.Any())
                    {
                        var list_Security = Exchange.ReadBSESecurity(BSESecurity[0].FullName, true);
                        var BSECMBhavcopy = Exchange.ReadBSECMBhavcopy(BsecmBhavPath[0].FullName,true);//to Fetch Closing Price

                        foreach (var _Security in list_Security)
                        {
                            long TotalTrades = 0;
                            double ClosingPrice = 0;
                            double PreviousClosPrice = 0;
                            var BsecmClosingDetails = BSECMBhavcopy.FirstOrDefault(v => v.Token == _Security.Token);//To fetch closing Price
                          
                            if (BsecmClosingDetails != null)
                            {
                                ClosingPrice = BsecmClosingDetails.Close; //Add a null check for closing price
                                PreviousClosPrice =BsecmClosingDetails.PreviousClose;
                                TotalTrades = BsecmClosingDetails.TotalTrades;
                            }


                            list_ContractMasterRows.Add($"({_Security.Token},'{_Security.Symbol}','EQ','{_Security.Series}','BSECM','{_Security.ScripName}'," +
                                $"'{_Security.CustomScripname}','EQ','{_Security.ExpiryUnix}',{0},{_Security.LotSize},{_Security.Token},'BSECM',{ClosingPrice},{PreviousClosPrice},'{_Security.ISIN}',{TotalTrades})");
                        }

                        sb_InsertCommand.Append(string.Join(",", list_ContractMasterRows));

                        using (MySqlConnection myConnToken = new MySqlConnection(_MySQLCon))
                        {
                            using (MySqlCommand myCmd = new MySqlCommand(sb_InsertCommand.ToString(), myConnToken))
                            {
                                myConnToken.Open();
                                myCmd.CommandType = CommandType.Text;
                                myCmd.ExecuteNonQuery();
                                myConnToken.Close();
                            }
                        }
                    }
                }
                catch (Exception ee)
                {
                    _logger.Error(ee, " : InsertTokensIntoDB BSECM");

                    AddToList("BSECM Tokens Upload failed. Please check logs for more details.", true);
                }


                #endregion

                #region MCX Contract

                list_ContractMasterRows.Clear();
                sb_InsertCommand.Clear();

                try
                {
                    string MCXFile = _MCXScripFile;//Filename

                    string MCXBhavFile = _MCXbhavcopy; //Bhavcopyname
                    string MCXBhavCopyPath = ((string)xmlDoc.Element("BOD-Utility").Element("MCX").Element("FILE").Element("BHAVCOPYPATH")).Trim(); //Bhavcopy path

                    //added Exists check on 27APR2021 by Amey   //MCXScrips.bcp
                    if (File.Exists(MCXFile))
                    {
                        sb_InsertCommand = new StringBuilder("INSERT IGNORE INTO tbl_contractmaster(Token,Symbol,InstrumentName,Series,Segment,ScripName,CustomScripName,ScripType,ExpiryUnix,StrikePrice,LotSize,Multiplier,UnderlyingToken,UnderlyingSegment,ClosingPrice) VALUES");

                        var list_MCXContract = Exchange.ReadMCXContract(MCXFile,isMcxContractFileContainsHeader); //MCXScrips.bcp
                        var MCXContracts = list_MCXContract.Where(v => v.ScripType == en_ScripType.XX).ToList();
                        var MCXBhavcopy = Exchange.ReadMCXBhavcopy(MCXBhavCopyPath + MCXBhavFile);//Added by Musharraf 22062023

                        foreach (var _Contract in list_MCXContract)
                        {
                            var USegment = "MCX";
                            var UnderlyingToken = -1;

                            if (_Contract.ScripType == en_ScripType.XX)
                                UnderlyingToken = _Contract.Token;
                            else
                            {
                                var temp = MCXContracts.Where(v => v.Symbol == _Contract.Symbol && v.Expiry.Month == _Contract.Expiry.Month && v.Expiry.Year == _Contract.Expiry.Year).FirstOrDefault();
                                if (temp != null)
                                    UnderlyingToken = temp.Token;
                                else
                                {
                                    //var twmp = list_EQSecurity.Where(v => v.Symbol == _Contract.Symbol).FirstOrDefault();
                                    //if (twmp != null)
                                    //{
                                    //    //UnderlyingToken = twmp.Token;
                                    //    //USegment = "NSECM";
                                    //    
                                    //}

                                    var temp1 = MCXContracts.Where(v => v.Symbol == _Contract.Symbol).OrderBy(v => v.Expiry).FirstOrDefault();

                                    if (temp1 != null)
                                        UnderlyingToken = temp1.Token;
                                }
                            }
                            var bhavcopyDetails = MCXBhavcopy.Where(v => v.CustomScripname == _Contract.CustomScripname).FirstOrDefault();//To fetch as per scripname
                            var ClosingPrice = (bhavcopyDetails != null) ? bhavcopyDetails.Close : 0;
                            //var scrip = (_Contract.ScripType == en_ScripType.XX);///Added for testing
                            //var details = MCXBhavcopy.Where(x => x.ScripType == en_ScripType.XX).FirstOrDefault();///Added for testing
                            if (UnderlyingToken != -1)
                                list_ContractMasterRows.Add($"({_Contract.Token},'{_Contract.Symbol}','{_Contract.Instrument}','-','MCX','{_Contract.ScripName}'," +
                                    $"'{_Contract.CustomScripname}','{_Contract.ScripType}','{_Contract.ExpiryUnix}',{_Contract.StrikePrice},{_Contract.LotSize},{_Contract.Multiplier}," +
                                    $"{UnderlyingToken},'{USegment}','{ClosingPrice}')");
                        }

                        sb_InsertCommand.Append(string.Join(",", list_ContractMasterRows));

                        using (MySqlConnection myConnToken = new MySqlConnection(_MySQLCon))
                        {
                            using (MySqlCommand myCmd = new MySqlCommand(sb_InsertCommand.ToString(), myConnToken))
                            {
                                myConnToken.Open();
                                myCmd.CommandType = CommandType.Text;
                                myCmd.ExecuteNonQuery();
                                myConnToken.Close();
                            }
                        }
                    }
                }
                catch (Exception ee)
                {
                    _logger.Error(ee, "InsertTokensIntoDB MCX : ");

                    AddToList("MCX Tokens Upload failed. Please check logs for more details.");
                }

                #endregion

                //added by omkar 
                #region BSEFO Security
                try
                {
                    list_ContractMasterRows.Clear();
                    sb_InsertCommand.Clear();

                    sb_InsertCommand = new StringBuilder("INSERT IGNORE INTO tbl_contractmaster(Token,Symbol,InstrumentName,Series,Segment,ScripName,CustomScripName,ScripType,ExpiryUnix,StrikePrice,LotSize,UnderlyingToken,UnderlyingSegment,ClosingPrice,SettlementPrice) VALUES");//Closing and Settlement added by Musharraf

                    DirectoryInfo _PrimeDirectory = new DirectoryInfo("C:\\Prime\\");
                    var BSEFOSecurity = _PrimeDirectory.GetFiles("BSE_EQD_CONTRACT_*.csv").OrderByDescending(v => v.LastWriteTime).ToList();


                    if (BSEFOSecurity.Any())
                    {
                        //var list2 = ReadBSEFOContract();//added for testing

                        var time = DateTime.Today;
                        var list_Security = Exchange.ReadBSEFOContract(BSEFOSecurity[0].FullName).Where(x => x.Expiry >= DateTime.Today).ToList();
                        var FutContracts = list_Security.Where(v => v.ScripType == NSEUtilitaire.en_ScripType.XX).ToList();

                        foreach (var _Security in list_Security)
                        {
                            var USegment = "BSEFO";
                            var UnderlyingToken = -1;

                            if (_Security.ScripType == NSEUtilitaire.en_ScripType.XX)
                                UnderlyingToken = _Security.Token;
                            else
                            {
                                var temp = FutContracts.Where(v => v.Symbol == _Security.Symbol && v.Expiry.Month == _Security.Expiry.Month && v.Expiry.Year == _Security.Expiry.Year).FirstOrDefault();
                                if (temp != null)
                                    UnderlyingToken = temp.Token;
                                else
                                {
                                    var twmp = list_EQSecurity.Where(v => v.Symbol == _Security.Symbol).FirstOrDefault();
                                    if (twmp != null)
                                    {
                                        UnderlyingToken = twmp.Token;
                                        USegment = "BSECM";
                                    }
                                }
                            }

                            //if (UnderlyingToken != -1)
                            //    list_ContractMasterRows.Add($"({_Security.Token},'{_Security.Symbol}','{_Security.Instrument}','-','BSEFO','{_Security.ScripName}'," +
                            //    $"'{_Security.CustomScripname}','{_Security.ScripType}','{_Security.ExpiryUnix}',{_Security.StrikePrice},{_Security.LotSize}," +
                            //    $"{UnderlyingToken},'BSEFO')");

                            list_ContractMasterRows.Add($"({_Security.Token},'{_Security.Symbol}','{_Security.Instrument}','-','BSEFO','{_Security.ScripName}'," +
                                    $"'{_Security.CustomScripname}','{_Security.ScripType}','{_Security.ExpiryUnix}',{_Security.StrikePrice},{_Security.LotSize}," +
                                    $"{UnderlyingToken},'{USegment}',{_Security.ClosePrice},{_Security.ClosePrice})");//pricefrombhavcopy.Close, pricefrombhavcopy.SettlePrice

                        }
                        sb_InsertCommand.Append(string.Join(",", list_ContractMasterRows));

                        using (MySqlConnection myConnToken = new MySqlConnection(_MySQLCon))
                        {
                            using (MySqlCommand myCmd = new MySqlCommand(sb_InsertCommand.ToString(), myConnToken))
                            {
                                myConnToken.Open();
                                myCmd.CommandType = CommandType.Text;
                                myCmd.ExecuteNonQuery();
                                myConnToken.Close();
                            }
                        }
                    }
                    else
                    {
                        _logger.Debug("BSEFO Contract file not found");
                    }
                }
                catch (Exception ee)
                {
                    _logger.Error(ee, " : InsertTokensIntoDB BSEFO");

                    AddToList("BSEFO Tokens Upload failed. Please check logs for more details.", true);
                }

                #endregion

                #region reading new files
                sb_InsertCommand.Clear();
                //APPSEC_COLLVAL_$date:ddMMyyyy$.csv added by Musharraf
                try
                {
                    sb_InsertCommand.Append("TRUNCATE tbl_collateralhaircut; ");
                    // Construct file path for the CSV file
                    string appSecCollValFile = collateralHaircut;
                    string csvFilePath = @"C:\Prime\Other\" + appSecCollValFile;

                    // Check if the CSV file exists
                    if (File.Exists(csvFilePath))
                    {
                        // StringBuilder to store the INSERT command
                        sb_InsertCommand.Append("INSERT IGNORE INTO tbl_collateralhaircut (Sr_No, SecurityName, ISIN, ClosingPrice, HairCut) VALUES ");

                        // Read the CSV file into a list of strings
                        List<string> csvLines = File.ReadAllLines(csvFilePath).ToList();

                        // Loop through each line in the CSV file
                        foreach (string line in csvLines)
                        {
                            if (line.Contains("Security"))
                            { continue; }
                            // Split the line into its individual values
                            string[] values = line.Split(',');

                            // Construct the INSERT command and add it to the StringBuilder
                            sb_InsertCommand.Append($"('{values[0]}', '{values[1]}', '{values[2]}', '{values[3]}', {values[4]}),");
                        }

                        // Remove the last comma from the INSERT command
                        sb_InsertCommand.Remove(sb_InsertCommand.Length - 1, 1);

                        // Create a new MySqlConnection object and open the connection
                        using (MySqlConnection conn = new MySqlConnection(_MySQLCon))
                        {
                            conn.Open();

                            // Create a new MySqlCommand object and set its properties
                            MySqlCommand cmd = new MySqlCommand(sb_InsertCommand.ToString(), conn);
                            cmd.CommandType = CommandType.Text;

                            // Execute the INSERT command
                            cmd.ExecuteNonQuery();

                            // Close the connection
                            conn.Close();
                        }
                    }
                    else
                    {
                        AddToList($"Collateral Haircut file Missing {appSecCollValFile},couldn't be uploaded in DB", true);
                    }
                }
                catch (Exception ex)
                {
                    _logger.Error(ex, " : InsertTokensIntoDB Collateral Haircut not Uploaded to DB check tbl_collateralhaircut");
                }

                sb_InsertCommand.Clear();
                //MF_VAR_$date:ddMMyyyy$.csv added by Musharraf
                try
                {
                    sb_InsertCommand.Append("TRUNCATE tbl_mfhaircut;");
                    // Construct file path for the CSV file
                    string mfhaircutfile = MFHaircut;
                    string[] saveFilePath = ds_Config.GET("SAVEPATH", "MF-HAIRCUT").SPL(',');
                    string csvFilePath = saveFilePath[0] + mfhaircutfile;

                    // Check if the CSV file exists
                    if (File.Exists(csvFilePath))
                    {
                        // StringBuilder to store the INSERT command
                        sb_InsertCommand.Append("INSERT IGNORE INTO tbl_mfhaircut (ISIN, SYMBOL, SERIES, TYPE, HAIRCUT, NAV) VALUES ");

                        // Read the CSV file into a list of strings
                        List<string> csvLines = File.ReadAllLines(csvFilePath).ToList();

                        // Loop through each line in the CSV file
                        foreach (string line in csvLines)
                        {
                            if (line.Equals("ISIN,SYMBOL,SERIES,TYPE,HAIRCUT,NAV"))
                            { continue; }
                            // Split the line into its individual values
                            string[] values = line.Split(',');

                            // Construct the INSERT command and add it to the StringBuilder
                            sb_InsertCommand.Append($"('{values[0]}', '{values[1]}', '{values[2]}', '{values[3]}', {values[4]}, {values[5]}),");
                        }

                        // Remove the last comma from the INSERT command
                        sb_InsertCommand.Remove(sb_InsertCommand.Length - 1, 1);

                        // Create a new MySqlConnection object and open the connection
                        using (MySqlConnection conn = new MySqlConnection(_MySQLCon))
                        {
                            conn.Open();

                            // Create a new MySqlCommand object and set its properties
                            MySqlCommand cmd = new MySqlCommand(sb_InsertCommand.ToString(), conn);
                            cmd.CommandType = CommandType.Text;

                            // Execute the INSERT command
                            cmd.ExecuteNonQuery();

                            // Close the connection
                            conn.Close();
                        }
                    }
                    else
                    {
                        AddToList($"MF Haircut file missing : {mfhaircutfile} couldn't be uploaded in DB", true);
                    }
                }
                catch (Exception ex)
                {
                    _logger.Error(ex, " : InsertTokensIntoDB MF Haircut not Uploaded to DB check tbl_mfhaircut");
                }
                sb_InsertCommand.Clear();
                //ind_close_all_$date:ddMMyyyy$.csv added by Musharraf
                try
                {
                    sb_InsertCommand.Append("TRUNCATE tbl_indexclosing; ");
                    // Construct file path for the CSV file
                    string SnapShotFile = DailySnapshot;
                    string csvFilePath = @"C:\Prime\" + SnapShotFile;

                    // Check if the CSV file exists
                    if (File.Exists(csvFilePath))
                    {
                        // StringBuilder to store the INSERT command
                        sb_InsertCommand.Append("INSERT IGNORE INTO tbl_indexclosing (IndexName, IndexDate, Open, High, Low, Closing, PointsChange, ChangePercent, Volume, TurnoverRsCr, PE, PB, DivYield) VALUES ");

                        // Read the CSV file into a list of strings
                        List<string> csvLines = File.ReadAllLines(csvFilePath).ToList();

                        // Loop through each line in the CSV file
                        foreach (string line in csvLines)
                        {
                            if (line.Contains("Index Name"))
                            {
                                continue;
                            }


                            // Split the line into its individual values
                            string[] values = line.Split(',');

                            // Replace empty or null values with 0.00 or 0
                            string openIndexValue = (string.IsNullOrEmpty(values[2]) || values[2].Equals("-")) ? "0.00" : values[2];
                            string highIndexValue = (string.IsNullOrEmpty(values[3]) || values[3].Equals("-")) ? "0.00" : values[3];
                            string lowIndexValue = (string.IsNullOrEmpty(values[4]) || values[4].Equals("-")) ? "0.00" : values[4];
                            string closingIndexValue = (string.IsNullOrEmpty(values[5]) || values[5].Equals("-")) ? "0.00" : values[5];
                            string pointsChange = (string.IsNullOrEmpty(values[6]) || values[6].Equals("-")) ? "0.00" : values[6];
                            string changePercent = (string.IsNullOrEmpty(values[7]) || values[7].Equals("-")) ? "0.00" : values[7];
                            string volume = (string.IsNullOrEmpty(values[8]) || values[8].Equals("-")) ? "0" : values[8];
                            string turnoverRsCr = (string.IsNullOrEmpty(values[9]) || values[9].Equals("-")) ? "0.00" : values[9];
                            string pe = (string.IsNullOrEmpty(values[10]) || values[10].Equals("-")) ? "0.00" : values[10];
                            string pb = (string.IsNullOrEmpty(values[11]) || values[11].Equals("-")) ? "0.00" : values[11];
                            string divYield = (string.IsNullOrEmpty(values[12]) || values[12].Equals("-")) ? "0.00" : values[12];

                            // Construct the INSERT command and add it to the StringBuilder
                            sb_InsertCommand.Append($"('{values[0]}', '{values[1]}', {openIndexValue}, {highIndexValue}, {lowIndexValue}, {closingIndexValue}, {pointsChange}, {changePercent}, {volume}, {turnoverRsCr}, {pe}, {pb}, {divYield}),");
                        }


                        // Remove the last comma from the INSERT command
                        sb_InsertCommand.Remove(sb_InsertCommand.Length - 1, 1);

                        // Create a new MySqlConnection object and open the connection
                        using (MySqlConnection conn = new MySqlConnection(_MySQLCon))
                        {
                            conn.Open();

                            // Create a new MySqlCommand object and set its properties
                            MySqlCommand cmd = new MySqlCommand(sb_InsertCommand.ToString(), conn);
                            cmd.CommandType = CommandType.Text;

                            // Execute the INSERT command
                            cmd.ExecuteNonQuery();

                            // Close the connection
                            conn.Close();
                        }
                    }
                    else
                    {
                        AddToList($"DailySnapShot file missing : {SnapShotFile} couldn't be uploaded in DB", true);
                    }
                }
                catch (Exception ex)
                {
                    _logger.Error(ex, " : InsertTokensIntoDB DailySnapShot not Uploaded in DB check tbl_indexclosing");
                }
                sb_InsertCommand.Clear();
                //fo_secban_$date:ddMMyyyy$.csv added by Musharraf 
                int CountofValues = 0; //To check the rows added in fo_secban
                try
                {

                   
                    sb_InsertCommand.Append("TRUNCATE tbl_fosecban;");
                    // Construct file path for the CSV file
                    string foSecban = FoSecban;
                    string[] saveFilepath = ds_Config.GET("SAVEPATH", "FO_SECBAN").SPL(',');
                    string _csvFilePath = saveFilepath[0] + foSecban;

                    var BanFile = new DirectoryInfo(saveFilepath[0]).GetFiles("fo_secban_*.csv").FirstOrDefault();

                    _logger.Debug("BanFilePath | " + saveFilepath[0]);

                    // Check if the CSV file exists
                    if (BanFile != null)
                    {
                        // StringBuilder to store the INSERT command
                        sb_InsertCommand.Append("INSERT IGNORE INTO tbl_fosecban (SrNo, ScripName) VALUES ");

                        // Read the CSV file into a list of strings
                        List<string> csvLines = File.ReadAllLines(BanFile.FullName).ToList();
                       
                        // Loop through each line in the CSV file
                        foreach (string line in csvLines)
                        {
                            // Split the line into its individual values
                            string[] values = line.Split(',');
                            if (values.Length > 1 && !values[0].Contains("Securities") || (values.Length > 1 ? !(string.IsNullOrWhiteSpace(values[1]) && string.IsNullOrEmpty(values[1])) : false))
                            {
                                // Construct the INSERT command and add it to the StringBuilder
                                sb_InsertCommand.Append($"({values[0]}, '{values[1]}'),");
                                CountofValues += 1;
                            }
                        }

                        // Remove the last comma from the INSERT command
                        sb_InsertCommand.Remove(sb_InsertCommand.Length - 1, 1);
                        var Query = "";
                        if (CountofValues > 0)
                        {
                            // Create a new MySqlConnection object and open the connection
                            Query = sb_InsertCommand.ToString();
                        }
                        else
                        {
                            
                            Query = "TRUNCATE tbl_fosecban;";
                            _logger.Debug("No scrips in BAN File for today");
                        }

                        using (MySqlConnection conn = new MySqlConnection(_MySQLCon))
                        {
                            conn.Open();

                            // Create a new MySqlCommand object and set its properties
                            MySqlCommand cmd = new MySqlCommand(Query, conn);
                            cmd.CommandType = CommandType.Text;

                            // Execute the INSERT command
                            cmd.ExecuteNonQuery();

                            // Close the connection
                            conn.Close();

                            _logger.Debug($"BAN File Rows added : {CountofValues}");
                        }
                    }
                    else
                    {
                        _logger.Debug("BanFile not fuund");
                        AddToList($"FOSecban file missing : {foSecban} couldn't be uploaded in DB", true);
                    }
                }
                catch (Exception ex)
                {
                    _logger.Error(ex, " : InsertTokensIntoDB FOSECBAN not Uploaded to DB check tbl_fosecban");
                    AddToList($"FOSecban file : {FoSecban} couldn't be uploaded in DB, check for File Data Format", true);
                }
                sb_InsertCommand.Clear();

                #endregion

                if (CMTokenUploaded == true && FOTokenUploaded == true)
                {
                    AddToList("Upload Token completed successfully");
                }
                else
                {
                    IsWorking = false;
                    btn_RestartAuto.Enabled = true;
                    AddToList("Upload Token unsuccessful", true);
                    SentMail("Upload Tokens");
                    if (list_ComponentStarted.Contains("FOUploadToken"))
                        list_ComponentStarted.Remove("FOUploadToken");
                    if (list_ComponentStarted.Contains("CMUploadToken"))
                        list_ComponentStarted.Remove("CMUploadToken");
                }
            }
            catch (Exception ee) { _logger.Error(ee); IsWorking = false; }
        }

        private void InsertTokensIntoDBUdiff()
        {
            AddToList("Upload Token Started");
            bool CMTokenUploaded = false; bool FOTokenUploaded = false;

            try
            {
                List<string> list_ContractMasterRows = new List<string>();
                List<Exchange.Security> list_EQSecurity = new List<Exchange.Security>();

                StringBuilder sb_InsertCommand = new StringBuilder("");

                #region IndexTokens and IndexScrip Upload
                //IndexTokens.csv
                try
                {
                    sb_InsertCommand.Append("TRUNCATE tbl_indextokens; ");
                    // Construct file path for the CSV file
                    string IndexTokens = "IndexTokens.csv";
                    string csvFilePath = @"C:\Prime\" + IndexTokens;

                    // Check if the CSV file exists
                    if (File.Exists(csvFilePath))
                    {
                        // StringBuilder to store the INSERT command
                        sb_InsertCommand.Append("INSERT IGNORE INTO tbl_indextokens (Symbol, Token, FullName, Segment) VALUES ");

                        // Read the CSV file into a list of strings
                        List<string> csvLines = File.ReadAllLines(csvFilePath).ToList();

                        // Loop through each line in the CSV file
                        foreach (string line in csvLines)
                        {
                            //if (line.Contains("Security"))
                            //{ continue; }
                            // Split the line into its individual values
                            string[] values = line.Split(',');

                            // Construct the INSERT command and add it to the StringBuilder
                            sb_InsertCommand.Append($"('{values[0]}', {values[1]}, '{values[2]}', '{values[3]}'),");
                        }

                        // Remove the last comma from the INSERT command
                        sb_InsertCommand.Remove(sb_InsertCommand.Length - 1, 1);

                        // Create a new MySqlConnection object and open the connection
                        using (MySqlConnection conn = new MySqlConnection(_MySQLCon))
                        {
                            conn.Open();

                            // Create a new MySqlCommand object and set its properties
                            MySqlCommand cmd = new MySqlCommand(sb_InsertCommand.ToString(), conn);
                            cmd.CommandType = CommandType.Text;

                            // Execute the INSERT command
                            cmd.ExecuteNonQuery();

                            // Close the connection
                            conn.Close();
                        }
                    }
                    else
                    {
                        AddToList($"IndexToken file Missing/Incorrect Format {IndexTokens} in {csvFilePath},couldn't be uploaded in DB", true);
                    }
                }
                catch (Exception ex)
                {
                    _logger.Error(ex, " : InsertTokensIntoDB IndexToken not Uploaded to DB check tbl_indextokens");
                }

                sb_InsertCommand.Clear();

                //IndexScrip.csv
                try
                {
                    sb_InsertCommand.Append("TRUNCATE tbl_indexscrip; ");
                    // Construct file path for the CSV file
                    string IndexScrip = "IndexScrip.csv";
                    string csvFilePath = @"C:\Prime\" + IndexScrip;

                    // Check if the CSV file exists
                    if (File.Exists(csvFilePath))
                    {
                        // StringBuilder to store the INSERT command
                        sb_InsertCommand.Append("INSERT IGNORE INTO tbl_indexscrip (FullName, Symbol) VALUES ");

                        // Read the CSV file into a list of strings
                        List<string> csvLines = File.ReadAllLines(csvFilePath).ToList();

                        // Loop through each line in the CSV file
                        foreach (string line in csvLines)
                        {
                            //if (line.Contains("Security"))
                            //{ continue; }
                            // Split the line into its individual values
                            string[] values = line.Split(',');

                            // Construct the INSERT command and add it to the StringBuilder
                            sb_InsertCommand.Append($"('{values[0]}', '{values[1]}'),");
                        }

                        // Remove the last comma from the INSERT command
                        sb_InsertCommand.Remove(sb_InsertCommand.Length - 1, 1);

                        // Create a new MySqlConnection object and open the connection
                        using (MySqlConnection conn = new MySqlConnection(_MySQLCon))
                        {
                            conn.Open();

                            // Create a new MySqlCommand object and set its properties
                            MySqlCommand cmd = new MySqlCommand(sb_InsertCommand.ToString(), conn);
                            cmd.CommandType = CommandType.Text;

                            // Execute the INSERT command
                            cmd.ExecuteNonQuery();

                            // Close the connection
                            conn.Close();
                        }
                    }
                    else
                    {
                        AddToList($"IndexScrip file Missing/Incorrect Format {IndexScrip} in {csvFilePath},couldn't be uploaded in DB", true);
                    }
                }
                catch (Exception ex)
                {
                    _logger.Error(ex, " : InsertTokensIntoDB IndexScrip not Uploaded to DB check tbl_indexscrip");
                }

                sb_InsertCommand.Clear();

                #endregion
                sb_InsertCommand.Append("TRUNCATE tbl_contractmaster; ");
                #region CM Security

                try
                {
                    //CM_security_fileName = "NSE_CM_security_22062023.csv";
                    string CMBhavCopyFile = _NSE_CM_bhavcopy;
                    var BhavcopyFile = new DirectoryInfo(@"C:\Prime\").GetFiles("BhavCopy_NSE_CM*.csv").OrderByDescending(x => x.LastWriteTime).FirstOrDefault();// @"C:\Prime\" + CMBhavCopyFile;
                    var Bhavcopy = new List<CMBhavcopy>();

                    if (BhavcopyFile != null)
                    {
                        Bhavcopy = Exchange.ReadCMBhavcopy(BhavcopyFile.FullName, true);//To fetch Closing and Settlement Price added by Musharra0f
                    }

                    //added Exists check on 27APR2021 by Amey       //security.txt
                    if (File.Exists("C:\\Prime\\" + CM_security_fileName))
                    {
                        sb_InsertCommand.Append("INSERT IGNORE INTO tbl_contractmaster(Token,Symbol,InstrumentName,Series,Segment,ScripName,CustomScripName,ScripType,ExpiryUnix,StrikePrice,LotSize,UnderlyingToken,UnderlyingSegment,ClosingPrice,SettlementPrice,Isin) VALUES");//,ClosingPrice,SettlementPrice added by Musharraf

                        var list_Security = Exchange.ReadSecurity("C:\\Prime\\" + CM_security_fileName, true);//security.txt
                        list_EQSecurity = list_Security.Where(v => v.Series == "EQ").ToList();

                        foreach (var _Security in list_Security)
                        {
                            //var bhavcopyDetails = Bhavcopy.FirstOrDefault(b => b.ScripName == _Security.ScripName && b.Series == _Security.Series);

                            var bhavcopyDetails = Bhavcopy.Where(v => v.CustomScripname == _Security.CustomScripname).FirstOrDefault();//To fetch as per scripname
                            var ClosingPrice = (bhavcopyDetails != null) ? bhavcopyDetails.Close : 0;

                            list_ContractMasterRows.Add($"({_Security.Token},'{_Security.Symbol}','EQ','{_Security.Series}','NSECM','{_Security.ScripName}'," +
                                $"'{_Security.CustomScripname}','EQ','{_Security.ExpiryUnix}',{0},{_Security.LotSize},{_Security.Token},'NSECM',{ClosingPrice},{0},'{_Security.ISIN}')");//bhavcopyDetails.Close

                            if (ClosingPrice <= 0)
                            {
                                _logger.Debug($"closing skipped for : {_Security.CustomScripname}");
                            }

                        }

                        ConcurrentDictionary<string, double> dict_IndexClosing = new ConcurrentDictionary<string, double>();
                        ConcurrentDictionary<string, string> dict_IndexScrips = new ConcurrentDictionary<string, string>();
                        ConcurrentDictionary<string, double> dict_bseIndexClosing = new ConcurrentDictionary<string, double>();
                        string SnapShotFile = DailySnapshot;
                        string IndCloseFilePath = @"C:\Prime\" + SnapShotFile;

                        try
                        {
                            var FilePath = "C://Prime//IndexScrip.csv";
                            if (File.Exists(FilePath))
                            {
                                string[] arr_Lines = File.ReadAllLines(FilePath);
                                foreach (var Line in arr_Lines)
                                {
                                    try
                                    {
                                        string[] arr_Fields = Line.Split(',');
                                        if (!dict_IndexScrips.ContainsKey(arr_Fields[0].Trim().ToUpper()))
                                            dict_IndexScrips.TryAdd(arr_Fields[0].Trim().ToUpper(), arr_Fields[1].Trim().ToUpper());
                                    }
                                    catch (Exception ee) { _logger.Error(ee); }
                                }
                            }
                        }
                        catch (Exception ee) { _logger.Error(ee); }

                        try
                        {
                            if (File.Exists(IndCloseFilePath))
                            {
                                foreach (var item in File.ReadAllLines(IndCloseFilePath))
                                {
                                    try
                                    {
                                        var arr_fields = item.Split(',');
                                        var IndexName = arr_fields[0].Trim().ToUpper();
                                        var closingPrice = Convert.ToDouble(arr_fields[5]);
                                        if (dict_IndexScrips.TryGetValue(IndexName, out string Symbol))
                                            dict_IndexClosing.TryAdd(Symbol, closingPrice);
                                    }
                                    catch (Exception ee) { _logger.Error(ee); }
                                }
                            }
                        }
                        catch (Exception ee) { _logger.Error(ee); }

                        try
                        {
                            var BseIndexCloseFile = new DirectoryInfo(@"C:\Prime\").GetFiles("index5_*").OrderByDescending(x => x.LastWriteTime).FirstOrDefault();
                            if (BseIndexCloseFile != null)
                            {
                                foreach (var line in File.ReadAllLines(BseIndexCloseFile.FullName))
                                {
                                    try
                                    {
                                        var arr_Fields = line.Split('|');
                                        var closing = Convert.ToDouble(arr_Fields[1]) / 100;
                                        dict_bseIndexClosing.TryAdd(arr_Fields[0], closing);
                                    }
                                    catch (Exception ee) { _logger.Error(ee); }

                                }
                            }
                        }
                        catch (Exception ee) { _logger.Error(ee); }

                        if (File.Exists("C:\\Prime\\IndexTokens.csv"))
                        {
                            foreach (var item in File.ReadAllLines("C:\\Prime\\IndexTokens.csv"))
                            {

                                string[] arr_Fields = item.Split(',');
                                string IndexName = arr_Fields[0];
                                string CustomScripName = $"{IndexName}|0|EQ|0";
                                double closingPrice = 0;
                                if (arr_Fields[3] == "BSE")
                                    dict_bseIndexClosing.TryGetValue(arr_Fields[4], out closingPrice);
                                else
                                    dict_IndexClosing.TryGetValue(IndexName, out closingPrice);

                                list_ContractMasterRows.Add($"({arr_Fields[1]},'{IndexName}','EQ','{en_InstrumentName.EQ}','{arr_Fields[3]}CM','{IndexName}-EQ'," +
                                      $"'{CustomScripName}','EQ',{0},{0},{1},{arr_Fields[1]},'{arr_Fields[3]}CM',{closingPrice},{0},'')");

                            }
                        }
                        else
                            XtraMessageBox.Show("Index Token file is not available.", "Error");

                        sb_InsertCommand.Append(string.Join(",", list_ContractMasterRows));

                        using (MySqlConnection myConnToken = new MySqlConnection(_MySQLCon))
                        {
                            using (MySqlCommand myCmd = new MySqlCommand(sb_InsertCommand.ToString(), myConnToken))
                            {
                                myConnToken.Open();
                                myCmd.CommandType = CommandType.Text;
                                myCmd.ExecuteNonQuery();
                                myConnToken.Close();
                            }
                        }
                        list_ComponentStarted.Add("CMUploadToken");
                        CMTokenUploaded = true;
                    }
                    else
                    {
                        AddToList("Security file not found", true);
                        CMTokenUploaded = false;
                    }
                }
                catch (Exception ee)
                {
                    _logger.Error(ee, " : InsertTokensIntoDB CM");

                    AddToList("CM Tokens Upload failed. Please check logs for more details.", true);

                    IsWorking = false;
                    btn_RestartAuto.Enabled = true;
                    btn_Settings.Enabled = true;
                    CMTokenUploaded = false;

                    if (list_ComponentStarted.Contains("CMUploadToken"))
                        list_ComponentStarted.Remove("CMUploadToken");
                    if (list_ComponentStarted.Contains("FOUploadToken"))
                        list_ComponentStarted.Remove("FOUploadToken");

                    SentMail("CM Tokens");

                }

                #endregion

                #region FO Contract

                list_ContractMasterRows.Clear();
                sb_InsertCommand.Clear();

                try
                {

                    //string FOBhavCopyFile = _FOBhavcopy.EndsWith(".zip") ? _FOBhavcopy.Substring(0, _FOBhavcopy.Length - 4) : _FOBhavcopy.EndsWith(".csv") ? _FOBhavcopy : throw new ArgumentException("Invalid file extension");
                    //string FilePath = FOBhavCopyFile.Contains("C:\\Prime\\") ? FOBhavCopyFile : (@"C:\Prime\" + FOBhavCopyFile);

                    //var Bhavcopy = Exchange.ReadFOBhavcopy(FilePath, true);//To fetch Closing and Settlement Price added by Musharraf
                    var BhavcopyFile = new DirectoryInfo(@"C:\Prime\").GetFiles("BhavCopy_NSE_FO*.csv").OrderByDescending(x => x.LastWriteTime).FirstOrDefault();// @"C:\Prime\" + CMBhavCopyFile;
                    var Bhavcopy = new List<FOBhavcopy>();

                    if (BhavcopyFile != null)
                    {
                        Bhavcopy = Exchange.ReadFOBhavcopy(BhavcopyFile.FullName,false,true);//To fetch Closing and Settlement Price added by Musharra0f
                    }


                    //added Exists check on 27APR2021 by Amey   //contract.txt
                    if (File.Exists("C:\\Prime\\" + FO_contract_fileName))
                    {
                        sb_InsertCommand = new StringBuilder("INSERT IGNORE INTO tbl_contractmaster(Token,Symbol,InstrumentName,Series,Segment,ScripName,CustomScripName,ScripType,ExpiryUnix,StrikePrice,LotSize,UnderlyingToken,UnderlyingSegment,ClosingPrice,SettlementPrice,OpenInterest) VALUES");//Closing and Settlement added by Musharraf

                        var list_Contract = Exchange.ReadContract("C:\\Prime\\" + FO_contract_fileName, true); //contract.txt
                        var FutContracts = list_Contract.Where(v => v.ScripType == en_ScripType.XX).ToList();

                        foreach (var _Contract in list_Contract)
                        {

                            var USegment = "NSEFO";
                            var UnderlyingToken = -1;

                            double ClosingPrice = 0;
                            double SettlingPrice = 0;
                            long OpenInterest = 0;

                            if (_Contract.ScripType == en_ScripType.XX)
                                UnderlyingToken = _Contract.Token;
                            else
                            {
                                var temp = FutContracts.Where(v => v.Symbol == _Contract.Symbol && v.Expiry.Month == _Contract.Expiry.Month && v.Expiry.Year == _Contract.Expiry.Year).FirstOrDefault();
                                if (temp != null)
                                    UnderlyingToken = temp.Token;
                                else
                                {
                                    //var twmp = list_EQSecurity.Where(v => v.Symbol == _Contract.Symbol).FirstOrDefault();
                                    //if (twmp != null)
                                    //{
                                    //    UnderlyingToken = twmp.Token;
                                    //    USegment = "NSECM";
                                    //}

                                    var temp1 = FutContracts.Where(v => v.Symbol == _Contract.Symbol).OrderByDescending(v => v.Expiry).FirstOrDefault();

                                    if (temp1 != null)
                                        UnderlyingToken = temp1.Token;

                                }
                            }

                            var pricefrombhavcopy = Bhavcopy.Where(v => v.CustomScripname == _Contract.CustomScripname).FirstOrDefault();//Added by Musharraf to fetch closing and settling price
                            if (pricefrombhavcopy != null)
                            {
                                ClosingPrice = pricefrombhavcopy.Close; //if null value is entered then Default should be zero
                                SettlingPrice = pricefrombhavcopy.SettlePrice;
                                OpenInterest = pricefrombhavcopy.OpenInterest;
                            }
                            //var ClosingPrice = (pricefrombhavcopy != null) ? pricefrombhavcopy.Close : 0; //if null value is entered then Default should be zero
                            //var SettlingPrice = (pricefrombhavcopy != null) ? pricefrombhavcopy.SettlePrice : 0;
                            //var OpenInterest = (pricefrombhavcopy != null) ? pricefrombhavcopy.OpenInterest : 0;
                            if (UnderlyingToken != -1)
                                list_ContractMasterRows.Add($"({_Contract.Token},'{_Contract.Symbol}','{_Contract.Instrument}','-','NSEFO','{_Contract.ScripName}'," +
                                    $"'{_Contract.CustomScripname}','{_Contract.ScripType}','{_Contract.ExpiryUnix}',{_Contract.StrikePrice},{_Contract.LotSize}," +
                                    $"{UnderlyingToken},'{USegment}',{ClosingPrice},{SettlingPrice},{OpenInterest})");//pricefrombhavcopy.Close, pricefrombhavcopy.SettlePrice
                        }

                        sb_InsertCommand.Append(string.Join(",", list_ContractMasterRows));

                        using (MySqlConnection myConnToken = new MySqlConnection(_MySQLCon))
                        {
                            using (MySqlCommand myCmd = new MySqlCommand(sb_InsertCommand.ToString(), myConnToken))
                            {
                                myConnToken.Open();
                                myCmd.CommandType = CommandType.Text;
                                myCmd.ExecuteNonQuery();
                                myConnToken.Close();
                            }
                        }

                        list_ComponentStarted.Add("FOUploadToken");
                        FOTokenUploaded = true;
                    }
                    else
                    {
                        AddToList("Contract file not found", true);
                        FOTokenUploaded = false;
                    }
                }
                catch (Exception ee)
                {
                    _logger.Error(ee, " : InsertTokensIntoDB FO");

                    AddToList("FO Tokens Upload failed. Please check logs for more details.", true);

                    IsWorking = false;
                    btn_RestartAuto.Enabled = true;
                    btn_Settings.Enabled = true;
                    FOTokenUploaded = false;

                    if (list_ComponentStarted.Contains("CMUploadToken"))
                        list_ComponentStarted.Remove("CMUploadToken");
                    if (list_ComponentStarted.Contains("FOUploadToken"))
                        list_ComponentStarted.Remove("FOUploadToken");


                    SentMail("FO Tokens");
                }

                #endregion

                #region CD Contract

                list_ContractMasterRows.Clear();
                sb_InsertCommand.Clear();

                try
                {

                    var BhavcopyFile = new DirectoryInfo(@"C:\Prime\").GetFiles("BhavCopy_NSE_CD*.csv").OrderByDescending(x => x.LastWriteTime).FirstOrDefault();// @"C:\Prime\" + CMBhavCopyFile;
                    var Bhavcopy = new List<CDBhavcopy>();

                    if (BhavcopyFile != null)
                    {
                        Bhavcopy = Exchange.ReadCDBhavcopy(BhavcopyFile.FullName, false, true);//To fetch Closing and Settlement Price added by Musharra0f
                    }

                    //added Exists check on 27APR2021 by Amey    //cd_contract.txt
                    if (File.Exists("C:\\Prime\\" + CD_contract_fileName))
                    {
                        sb_InsertCommand = new StringBuilder("INSERT IGNORE INTO tbl_contractmaster(Token,Symbol,InstrumentName,Series,Segment,ScripName,CustomScripName,ScripType,ExpiryUnix,StrikePrice,LotSize,UnderlyingToken,UnderlyingSegment,ClosingPrice,SettlementPrice,Multiplier,PrevClosingPrice,CD_OpenInterest) VALUES");//Closing and Settlement added by Musharraf

                        var list_CDContract = Exchange.ReadCDContract("C:\\Prime\\" + CD_contract_fileName, true);//cd_contract.txt
                        var FutContracts = list_CDContract.Where(v => v.ScripType == en_ScripType.XX).ToList();

                        foreach (var _CDContract in list_CDContract)
                        {
                            var UnderlyingToken = -1;
                            if (_CDContract.ScripType == en_ScripType.XX)
                                UnderlyingToken = _CDContract.Token;
                            else
                            {
                                var temp = FutContracts.Where(v => v.Symbol == _CDContract.Symbol && v.Expiry.Month == _CDContract.Expiry.Month && v.Expiry.Year == _CDContract.Expiry.Year).FirstOrDefault();
                                if (temp != null)
                                    UnderlyingToken = temp.Token;
                            }

                            var cdbhavcopydetails = Bhavcopy.FirstOrDefault(v => v.Symbol == _CDContract.Symbol && v.Expiry.Date == _CDContract.Expiry.Date && v.StrikePrice == _CDContract.StrikePrice && v.ScripType == _CDContract.ScripType);//To fetch closing Price
                            var ClosingPrice = (cdbhavcopydetails == null) ? 0 : cdbhavcopydetails.Close; //Add a null check for closing price
                            var PreviousClosPrice = (cdbhavcopydetails == null) ? 0 : cdbhavcopydetails.PreviousClose;
                            var CD_OpenInterest = (cdbhavcopydetails == null) ? 0 : cdbhavcopydetails.OpenInterest;

                            if (UnderlyingToken != -1)
                                list_ContractMasterRows.Add($"({_CDContract.Token},'{_CDContract.Symbol}','{_CDContract.Instrument}','-','NSECD','{_CDContract.ScripName}'," +
                                    $"'{_CDContract.CustomScripname}','{_CDContract.ScripType}','{_CDContract.ExpiryUnix}',{_CDContract.StrikePrice},{_CDContract.LotSize}," +
                                    $"{UnderlyingToken},'NSECD',{ClosingPrice},{0},{_CDContract.Multiplier},{PreviousClosPrice},{CD_OpenInterest})");
                        }

                        sb_InsertCommand.Append(string.Join(",", list_ContractMasterRows));

                        using (MySqlConnection myConnToken = new MySqlConnection(_MySQLCon))
                        {
                            using (MySqlCommand myCmd = new MySqlCommand(sb_InsertCommand.ToString(), myConnToken))
                            {
                                myConnToken.Open();
                                myCmd.CommandType = CommandType.Text;
                                myCmd.ExecuteNonQuery();
                                myConnToken.Close();
                            }
                        }
                    }
                }
                catch (Exception ee)
                {
                    _logger.Error(ee, " : InsertTokensIntoDB CD");

                    AddToList("CD Tokens Upload failed. Please check logs for more details.", true);
                }

                #endregion

                #region BSECM Security

                try
                {
                    list_ContractMasterRows.Clear();
                    sb_InsertCommand.Clear();
                    sb_InsertCommand.Append("INSERT IGNORE INTO tbl_contractmaster(Token,Symbol,InstrumentName,Series,Segment,ScripName,CustomScripName,ScripType,ExpiryUnix,StrikePrice,LotSize,UnderlyingToken,UnderlyingSegment,ClosingPrice,PrevClosingPrice,Isin,TotalTrades) VALUES");

                    DirectoryInfo _PrimeDirectory = new DirectoryInfo("C:\\Prime\\");
                    var BSESecurity = _PrimeDirectory.GetFiles("BSE_EQ_SCRIP_*.csv").OrderByDescending(v => v.LastWriteTime).ToList();//changed .txt to .csv by Musharraf 21st April 2023

                    var BsecmBhavPath = _PrimeDirectory.GetFiles("BhavCopy_BSE_CM_*.csv").OrderByDescending(v => v.LastWriteTime).ToList();

                    if (BSESecurity.Any())
                    {
                        var list_Security = Exchange.ReadBSESecurity(BSESecurity[0].FullName, true);
                        var BSECMBhavcopy = Exchange.ReadBSECMBhavcopy(BsecmBhavPath[0].FullName, true);//to Fetch Closing Price

                        foreach (var _Security in list_Security)
                        {
                            long TotalTrades = 0;
                            double ClosingPrice = 0;
                            double PreviousClosPrice = 0;
                            var BsecmClosingDetails = BSECMBhavcopy.FirstOrDefault(v => v.Token == _Security.Token);//To fetch closing Price

                            if (BsecmClosingDetails != null)
                            {
                                ClosingPrice = BsecmClosingDetails.Close; //Add a null check for closing price
                                PreviousClosPrice = BsecmClosingDetails.PreviousClose;
                                TotalTrades = BsecmClosingDetails.TotalTrades;
                            }


                            list_ContractMasterRows.Add($"({_Security.Token},'{_Security.Symbol}','EQ','{_Security.Series}','BSECM','{_Security.ScripName}'," +
                                $"'{_Security.CustomScripname}','EQ','{_Security.ExpiryUnix}',{0},{_Security.LotSize},{_Security.Token},'BSECM',{ClosingPrice},{PreviousClosPrice},'{_Security.ISIN}',{TotalTrades})");
                        }

                        sb_InsertCommand.Append(string.Join(",", list_ContractMasterRows));

                        using (MySqlConnection myConnToken = new MySqlConnection(_MySQLCon))
                        {
                            using (MySqlCommand myCmd = new MySqlCommand(sb_InsertCommand.ToString(), myConnToken))
                            {
                                myConnToken.Open();
                                myCmd.CommandType = CommandType.Text;
                                myCmd.ExecuteNonQuery();
                                myConnToken.Close();
                            }
                        }
                    }
                }
                catch (Exception ee)
                {
                    _logger.Error(ee, " : InsertTokensIntoDB BSECM");

                    AddToList("BSECM Tokens Upload failed. Please check logs for more details.", true);
                }


                #endregion

                #region MCX Contract

                list_ContractMasterRows.Clear();
                sb_InsertCommand.Clear();

                try
                {
                    string MCXFile = _MCXScripFile;//Filename

                    string MCXBhavFile = _MCXbhavcopy; //Bhavcopyname
                    string MCXBhavCopyPath = ((string)xmlDoc.Element("BOD-Utility").Element("MCX").Element("FILE").Element("BHAVCOPYPATH")).Trim(); //Bhavcopy path

                    var BhavcopyFile = new DirectoryInfo(MCXBhavCopyPath).GetFiles("BhavCopy_MCXCCL_CO*.csv").OrderByDescending(x => x.LastWriteTime).FirstOrDefault();// @"C:\Prime\" + CMBhavCopyFile;
                    var MCXBhavcopy = new List<MCXBhavcopy>();

                    if (BhavcopyFile != null)
                    {
                        MCXBhavcopy = Exchange.ReadMCXBhavcopy(BhavcopyFile.FullName,true);//To fetch Closing and Settlement Price added by Musharra0f
                    }

                    //added Exists check on 27APR2021 by Amey   //MCXScrips.bcp
                    if (File.Exists(MCXFile))
                    {
                        sb_InsertCommand = new StringBuilder("INSERT IGNORE INTO tbl_contractmaster(Token,Symbol,InstrumentName,Series,Segment,ScripName,CustomScripName,ScripType,ExpiryUnix,StrikePrice,LotSize,Multiplier,UnderlyingToken,UnderlyingSegment,ClosingPrice) VALUES");

                        var list_MCXContract = Exchange.ReadMCXContract(MCXFile,isMcxContractFileContainsHeader); //MCXScrips.bcp
                        var MCXContracts = list_MCXContract.Where(v => v.ScripType == en_ScripType.XX).ToList();
                        //var MCXBhavcopy = Exchange.ReadMCXBhavcopy(MCXBhavCopyPath + MCXBhavFile);//Added by Musharraf 22062023

                        foreach (var _Contract in list_MCXContract)
                        {
                            var USegment = "MCX";
                            var UnderlyingToken = -1;

                            if (_Contract.ScripType == en_ScripType.XX)
                                UnderlyingToken = _Contract.Token;
                            else
                            {
                                var temp = MCXContracts.Where(v => v.Symbol == _Contract.Symbol && v.Expiry.Month == _Contract.Expiry.Month && v.Expiry.Year == _Contract.Expiry.Year).FirstOrDefault();
                                if (temp != null)
                                    UnderlyingToken = temp.Token;
                                else
                                {
                                    //var twmp = list_EQSecurity.Where(v => v.Symbol == _Contract.Symbol).FirstOrDefault();
                                    //if (twmp != null)
                                    //{
                                    //    //UnderlyingToken = twmp.Token;
                                    //    //USegment = "NSECM";
                                    //    
                                    //}

                                    var temp1 = MCXContracts.Where(v => v.Symbol == _Contract.Symbol).OrderBy(v => v.Expiry).FirstOrDefault();

                                    if (temp1 != null)
                                        UnderlyingToken = temp1.Token;
                                }
                            }
                            var bhavcopyDetails = MCXBhavcopy.Where(v => v.CustomScripname == _Contract.CustomScripname).FirstOrDefault();//To fetch as per scripname
                            var ClosingPrice = (bhavcopyDetails != null) ? bhavcopyDetails.Close : 0;
                            //var scrip = (_Contract.ScripType == en_ScripType.XX);///Added for testing
                            //var details = MCXBhavcopy.Where(x => x.ScripType == en_ScripType.XX).FirstOrDefault();///Added for testing
                            if (UnderlyingToken != -1)
                                list_ContractMasterRows.Add($"({_Contract.Token},'{_Contract.Symbol}','{_Contract.Instrument}','-','MCX','{_Contract.ScripName}'," +
                                    $"'{_Contract.CustomScripname}','{_Contract.ScripType}','{_Contract.ExpiryUnix}',{_Contract.StrikePrice},{_Contract.LotSize},{_Contract.Multiplier}," +
                                    $"{UnderlyingToken},'{USegment}','{ClosingPrice}')");
                        }

                        sb_InsertCommand.Append(string.Join(",", list_ContractMasterRows));

                        using (MySqlConnection myConnToken = new MySqlConnection(_MySQLCon))
                        {
                            using (MySqlCommand myCmd = new MySqlCommand(sb_InsertCommand.ToString(), myConnToken))
                            {
                                myConnToken.Open();
                                myCmd.CommandType = CommandType.Text;
                                myCmd.ExecuteNonQuery();
                                myConnToken.Close();
                            }
                        }
                    }
                }
                catch (Exception ee)
                {
                    _logger.Error(ee, "InsertTokensIntoDB MCX : ");

                    AddToList("MCX Tokens Upload failed. Please check logs for more details.");
                }

                #endregion

                //added by omkar 
                #region BSEFO Security
                try
                {
                    list_ContractMasterRows.Clear();
                    sb_InsertCommand.Clear();

                    sb_InsertCommand = new StringBuilder("INSERT IGNORE INTO tbl_contractmaster(Token,Symbol,InstrumentName,Series,Segment,ScripName,CustomScripName,ScripType,ExpiryUnix,StrikePrice,LotSize,UnderlyingToken,UnderlyingSegment,ClosingPrice,SettlementPrice) VALUES");//Closing and Settlement added by Musharraf

                    DirectoryInfo _PrimeDirectory = new DirectoryInfo("C:\\Prime\\");
                    var BSEFOSecurity = _PrimeDirectory.GetFiles("BSE_EQD_CONTRACT_*.csv").OrderByDescending(v => v.LastWriteTime).ToList();


                    if (BSEFOSecurity.Any())
                    {
                        //var list2 = ReadBSEFOContract();//added for testing

                        var time = DateTime.Today;
                        var list_Security = Exchange.ReadBSEFOContract(BSEFOSecurity[0].FullName).Where(x => x.Expiry >= DateTime.Today).ToList();
                        var FutContracts = list_Security.Where(v => v.ScripType == NSEUtilitaire.en_ScripType.XX).ToList();

                        foreach (var _Security in list_Security)
                        {
                            var USegment = "BSEFO";
                            var UnderlyingToken = -1;

                            if (_Security.ScripType == NSEUtilitaire.en_ScripType.XX)
                                UnderlyingToken = _Security.Token;
                            else
                            {
                                var temp = FutContracts.Where(v => v.Symbol == _Security.Symbol && v.Expiry.Month == _Security.Expiry.Month && v.Expiry.Year == _Security.Expiry.Year).FirstOrDefault();
                                if (temp != null)
                                    UnderlyingToken = temp.Token;
                                else
                                {
                                    var twmp = list_EQSecurity.Where(v => v.Symbol == _Security.Symbol).FirstOrDefault();
                                    if (twmp != null)
                                    {
                                        UnderlyingToken = twmp.Token;
                                        USegment = "BSECM";
                                    }
                                }
                            }

                            //if (UnderlyingToken != -1)
                            //    list_ContractMasterRows.Add($"({_Security.Token},'{_Security.Symbol}','{_Security.Instrument}','-','BSEFO','{_Security.ScripName}'," +
                            //    $"'{_Security.CustomScripname}','{_Security.ScripType}','{_Security.ExpiryUnix}',{_Security.StrikePrice},{_Security.LotSize}," +
                            //    $"{UnderlyingToken},'BSEFO')");

                            list_ContractMasterRows.Add($"({_Security.Token},'{_Security.Symbol}','{_Security.Instrument}','-','BSEFO','{_Security.ScripName}'," +
                                    $"'{_Security.CustomScripname}','{_Security.ScripType}','{_Security.ExpiryUnix}',{_Security.StrikePrice},{_Security.LotSize}," +
                                    $"{UnderlyingToken},'{USegment}',{_Security.ClosePrice},{_Security.ClosePrice})");//pricefrombhavcopy.Close, pricefrombhavcopy.SettlePrice

                        }
                        sb_InsertCommand.Append(string.Join(",", list_ContractMasterRows));

                        using (MySqlConnection myConnToken = new MySqlConnection(_MySQLCon))
                        {
                            using (MySqlCommand myCmd = new MySqlCommand(sb_InsertCommand.ToString(), myConnToken))
                            {
                                myConnToken.Open();
                                myCmd.CommandType = CommandType.Text;
                                myCmd.ExecuteNonQuery();
                                myConnToken.Close();
                            }
                        }
                    }
                    else
                    {
                        _logger.Debug("BSEFO Contract file not found");
                    }
                }
                catch (Exception ee)
                {
                    _logger.Error(ee, " : InsertTokensIntoDB BSEFO");

                    AddToList("BSEFO Tokens Upload failed. Please check logs for more details.", true);
                }

                #endregion

                #region reading new files
                sb_InsertCommand.Clear();
                //APPSEC_COLLVAL_$date:ddMMyyyy$.csv added by Musharraf
                try
                {
                    sb_InsertCommand.Append("TRUNCATE tbl_collateralhaircut; ");
                    // Construct file path for the CSV file
                    string appSecCollValFile = collateralHaircut;
                    string csvFilePath = @"C:\Prime\Other\" + appSecCollValFile;

                    // Check if the CSV file exists
                    if (File.Exists(csvFilePath))
                    {
                        // StringBuilder to store the INSERT command
                        sb_InsertCommand.Append("INSERT IGNORE INTO tbl_collateralhaircut (Sr_No, SecurityName, ISIN, ClosingPrice, HairCut) VALUES ");

                        // Read the CSV file into a list of strings
                        List<string> csvLines = File.ReadAllLines(csvFilePath).ToList();

                        // Loop through each line in the CSV file
                        foreach (string line in csvLines)
                        {
                            if (line.Contains("Security"))
                            { continue; }
                            // Split the line into its individual values
                            string[] values = line.Split(',');

                            // Construct the INSERT command and add it to the StringBuilder
                            sb_InsertCommand.Append($"('{values[0]}', '{values[1]}', '{values[2]}', '{values[3]}', {values[4]}),");
                        }

                        // Remove the last comma from the INSERT command
                        sb_InsertCommand.Remove(sb_InsertCommand.Length - 1, 1);

                        // Create a new MySqlConnection object and open the connection
                        using (MySqlConnection conn = new MySqlConnection(_MySQLCon))
                        {
                            conn.Open();

                            // Create a new MySqlCommand object and set its properties
                            MySqlCommand cmd = new MySqlCommand(sb_InsertCommand.ToString(), conn);
                            cmd.CommandType = CommandType.Text;

                            // Execute the INSERT command
                            cmd.ExecuteNonQuery();

                            // Close the connection
                            conn.Close();
                        }
                    }
                    else
                    {
                        AddToList($"Collateral Haircut file Missing {appSecCollValFile},couldn't be uploaded in DB", true);
                    }
                }
                catch (Exception ex)
                {
                    _logger.Error(ex, " : InsertTokensIntoDB Collateral Haircut not Uploaded to DB check tbl_collateralhaircut");
                }

                sb_InsertCommand.Clear();
                //MF_VAR_$date:ddMMyyyy$.csv added by Musharraf
                try
                {
                    sb_InsertCommand.Append("TRUNCATE tbl_mfhaircut;");
                    // Construct file path for the CSV file
                    string mfhaircutfile = MFHaircut;
                    string[] saveFilePath = ds_Config.GET("SAVEPATH", "MF-HAIRCUT").SPL(',');
                    string csvFilePath = saveFilePath[0] + mfhaircutfile;

                    // Check if the CSV file exists
                    if (File.Exists(csvFilePath))
                    {
                        // StringBuilder to store the INSERT command
                        sb_InsertCommand.Append("INSERT IGNORE INTO tbl_mfhaircut (ISIN, SYMBOL, SERIES, TYPE, HAIRCUT, NAV) VALUES ");

                        // Read the CSV file into a list of strings
                        List<string> csvLines = File.ReadAllLines(csvFilePath).ToList();

                        // Loop through each line in the CSV file
                        foreach (string line in csvLines)
                        {
                            if (line.Equals("ISIN,SYMBOL,SERIES,TYPE,HAIRCUT,NAV"))
                            { continue; }
                            // Split the line into its individual values
                            string[] values = line.Split(',');

                            // Construct the INSERT command and add it to the StringBuilder
                            sb_InsertCommand.Append($"('{values[0]}', '{values[1]}', '{values[2]}', '{values[3]}', {values[4]}, {values[5]}),");
                        }

                        // Remove the last comma from the INSERT command
                        sb_InsertCommand.Remove(sb_InsertCommand.Length - 1, 1);

                        // Create a new MySqlConnection object and open the connection
                        using (MySqlConnection conn = new MySqlConnection(_MySQLCon))
                        {
                            conn.Open();

                            // Create a new MySqlCommand object and set its properties
                            MySqlCommand cmd = new MySqlCommand(sb_InsertCommand.ToString(), conn);
                            cmd.CommandType = CommandType.Text;

                            // Execute the INSERT command
                            cmd.ExecuteNonQuery();

                            // Close the connection
                            conn.Close();
                        }
                    }
                    else
                    {
                        AddToList($"MF Haircut file missing : {mfhaircutfile} couldn't be uploaded in DB", true);
                    }
                }
                catch (Exception ex)
                {
                    _logger.Error(ex, " : InsertTokensIntoDB MF Haircut not Uploaded to DB check tbl_mfhaircut");
                }
                sb_InsertCommand.Clear();
                //ind_close_all_$date:ddMMyyyy$.csv added by Musharraf
                try
                {
                    sb_InsertCommand.Append("TRUNCATE tbl_indexclosing; ");
                    // Construct file path for the CSV file
                    string SnapShotFile = DailySnapshot;
                    string csvFilePath = @"C:\Prime\" + SnapShotFile;

                    // Check if the CSV file exists
                    if (File.Exists(csvFilePath))
                    {
                        // StringBuilder to store the INSERT command
                        sb_InsertCommand.Append("INSERT IGNORE INTO tbl_indexclosing (IndexName, IndexDate, Open, High, Low, Closing, PointsChange, ChangePercent, Volume, TurnoverRsCr, PE, PB, DivYield) VALUES ");

                        // Read the CSV file into a list of strings
                        List<string> csvLines = File.ReadAllLines(csvFilePath).ToList();

                        // Loop through each line in the CSV file
                        foreach (string line in csvLines)
                        {
                            if (line.Contains("Index Name"))
                            {
                                continue;
                            }


                            // Split the line into its individual values
                            string[] values = line.Split(',');

                            // Replace empty or null values with 0.00 or 0
                            string openIndexValue = (string.IsNullOrEmpty(values[2]) || values[2].Equals("-")) ? "0.00" : values[2];
                            string highIndexValue = (string.IsNullOrEmpty(values[3]) || values[3].Equals("-")) ? "0.00" : values[3];
                            string lowIndexValue = (string.IsNullOrEmpty(values[4]) || values[4].Equals("-")) ? "0.00" : values[4];
                            string closingIndexValue = (string.IsNullOrEmpty(values[5]) || values[5].Equals("-")) ? "0.00" : values[5];
                            string pointsChange = (string.IsNullOrEmpty(values[6]) || values[6].Equals("-")) ? "0.00" : values[6];
                            string changePercent = (string.IsNullOrEmpty(values[7]) || values[7].Equals("-")) ? "0.00" : values[7];
                            string volume = (string.IsNullOrEmpty(values[8]) || values[8].Equals("-")) ? "0" : values[8];
                            string turnoverRsCr = (string.IsNullOrEmpty(values[9]) || values[9].Equals("-")) ? "0.00" : values[9];
                            string pe = (string.IsNullOrEmpty(values[10]) || values[10].Equals("-")) ? "0.00" : values[10];
                            string pb = (string.IsNullOrEmpty(values[11]) || values[11].Equals("-")) ? "0.00" : values[11];
                            string divYield = (string.IsNullOrEmpty(values[12]) || values[12].Equals("-")) ? "0.00" : values[12];

                            // Construct the INSERT command and add it to the StringBuilder
                            sb_InsertCommand.Append($"('{values[0]}', '{values[1]}', {openIndexValue}, {highIndexValue}, {lowIndexValue}, {closingIndexValue}, {pointsChange}, {changePercent}, {volume}, {turnoverRsCr}, {pe}, {pb}, {divYield}),");
                        }


                        // Remove the last comma from the INSERT command
                        sb_InsertCommand.Remove(sb_InsertCommand.Length - 1, 1);

                        // Create a new MySqlConnection object and open the connection
                        using (MySqlConnection conn = new MySqlConnection(_MySQLCon))
                        {
                            conn.Open();

                            // Create a new MySqlCommand object and set its properties
                            MySqlCommand cmd = new MySqlCommand(sb_InsertCommand.ToString(), conn);
                            cmd.CommandType = CommandType.Text;

                            // Execute the INSERT command
                            cmd.ExecuteNonQuery();

                            // Close the connection
                            conn.Close();
                        }
                    }
                    else
                    {
                        AddToList($"DailySnapShot file missing : {SnapShotFile} couldn't be uploaded in DB", true);
                    }
                }
                catch (Exception ex)
                {
                    _logger.Error(ex, " : InsertTokensIntoDB DailySnapShot not Uploaded in DB check tbl_indexclosing");
                }
                sb_InsertCommand.Clear();
                //fo_secban_$date:ddMMyyyy$.csv added by Musharraf 
                int CountofValues = 0; //To check the rows added in fo_secban
                try
                {


                    sb_InsertCommand.Append("TRUNCATE tbl_fosecban;");
                    // Construct file path for the CSV file
                    string foSecban = FoSecban;
                    string[] saveFilepath = ds_Config.GET("SAVEPATH", "FO_SECBAN").SPL(',');
                    string _csvFilePath = saveFilepath[0] + foSecban;

                    var BanFile = new DirectoryInfo(saveFilepath[0]).GetFiles("fo_secban_*.csv").FirstOrDefault();

                    _logger.Debug("BanFilePath | " + saveFilepath[0]);

                    // Check if the CSV file exists
                    if (BanFile != null)
                    {
                        // StringBuilder to store the INSERT command
                        sb_InsertCommand.Append("INSERT IGNORE INTO tbl_fosecban (SrNo, ScripName) VALUES ");

                        // Read the CSV file into a list of strings
                        List<string> csvLines = File.ReadAllLines(BanFile.FullName).ToList();

                        // Loop through each line in the CSV file
                        foreach (string line in csvLines)
                        {
                            // Split the line into its individual values
                            string[] values = line.Split(',');
                            if (values.Length > 1 && !values[0].Contains("Securities") || (values.Length > 1 ? !(string.IsNullOrWhiteSpace(values[1]) && string.IsNullOrEmpty(values[1])) : false))
                            {
                                // Construct the INSERT command and add it to the StringBuilder
                                sb_InsertCommand.Append($"({values[0]}, '{values[1]}'),");
                                CountofValues += 1;
                            }
                        }

                        // Remove the last comma from the INSERT command
                        sb_InsertCommand.Remove(sb_InsertCommand.Length - 1, 1);
                        var Query = "";
                        if (CountofValues > 0)
                        {
                            // Create a new MySqlConnection object and open the connection
                            Query = sb_InsertCommand.ToString();
                        }
                        else
                        {

                            Query = "TRUNCATE tbl_fosecban;";
                            _logger.Debug("No scrips in BAN File for today");
                        }

                        using (MySqlConnection conn = new MySqlConnection(_MySQLCon))
                        {
                            conn.Open();

                            // Create a new MySqlCommand object and set its properties
                            MySqlCommand cmd = new MySqlCommand(Query, conn);
                            cmd.CommandType = CommandType.Text;

                            // Execute the INSERT command
                            cmd.ExecuteNonQuery();

                            // Close the connection
                            conn.Close();

                            _logger.Debug($"BAN File Rows added : {CountofValues}");
                        }
                    }
                    else
                    {
                        _logger.Debug("BanFile not fuund");
                        AddToList($"FOSecban file missing : {foSecban} couldn't be uploaded in DB", true);
                    }
                }
                catch (Exception ex)
                {
                    _logger.Error(ex, " : InsertTokensIntoDB FOSECBAN not Uploaded to DB check tbl_fosecban");
                    AddToList($"FOSecban file : {FoSecban} couldn't be uploaded in DB, check for File Data Format", true);
                }
                sb_InsertCommand.Clear();

                #endregion

                if (CMTokenUploaded == true && FOTokenUploaded == true)
                {
                    AddToList("Upload Token completed successfully");
                }
                else
                {
                    IsWorking = false;
                    btn_RestartAuto.Enabled = true;
                    AddToList("Upload Token unsuccessful", true);
                    SentMail("Upload Tokens");
                    if (list_ComponentStarted.Contains("FOUploadToken"))
                        list_ComponentStarted.Remove("FOUploadToken");
                    if (list_ComponentStarted.Contains("CMUploadToken"))
                        list_ComponentStarted.Remove("CMUploadToken");
                }
            }
            catch (Exception ee) { _logger.Error(ee); IsWorking = false; }
        }



        private void _InsertTokensIntoDB()
        {

            #region FO Contract
            var list_ContractMasterRows = new List<string>();
            list_ContractMasterRows.Clear();
            StringBuilder sb_InsertCommand = new StringBuilder();
            sb_InsertCommand.Clear();

            try
            {

                //string FOBhavCopyFile = _FOBhavcopy.EndsWith(".zip") ? _FOBhavcopy.Substring(0, _FOBhavcopy.Length - 4) : _FOBhavcopy.EndsWith(".csv") ? _FOBhavcopy : throw new ArgumentException("Invalid file extension");
                //string FilePath = FOBhavCopyFile.Contains("C:\\Prime\\") ? FOBhavCopyFile : (@"C:\Prime\" + FOBhavCopyFile);

                var Bhavcopy = Exchange.ReadFOBhavcopy("C:\\Prime\\NSE_FO_bhavcopy_23022024.csv", true);//To fetch Closing and Settlement Price added by Musharraf

                FO_contract_fileName = "NSE_FO_contract_22022024.csv";

                //added Exists check on 27APR2021 by Amey   //contract.txt
                if (File.Exists("C:\\Prime\\" + FO_contract_fileName))
                {
                    sb_InsertCommand = new StringBuilder("INSERT IGNORE INTO tbl_contractmaster(Token,Symbol,InstrumentName,Series,Segment,ScripName,CustomScripName,ScripType,ExpiryUnix,StrikePrice,LotSize,UnderlyingToken,UnderlyingSegment,ClosingPrice,SettlementPrice,OpenInterest) VALUES");//Closing and Settlement added by Musharraf

                    var list_Contract = Exchange.ReadContract("C:\\Prime\\" + FO_contract_fileName, true).Where(X=>X.Symbol=="IDEA"); //contract.txt
                    var FutContracts = list_Contract.Where(v => v.ScripType == en_ScripType.XX).ToList();

                    foreach (var _Contract in list_Contract)
                    {
                        //if(_Contract.Symbol == "IDEA")
                        //{

                        //}


                        var USegment = "NSEFO";
                        var UnderlyingToken = -1;

                        double ClosingPrice = 0;
                        double SettlingPrice = 0;
                        long OpenInterest = 0;

                        if (_Contract.ScripType == en_ScripType.XX)
                            UnderlyingToken = _Contract.Token;
                        else
                        {
                            var temp = FutContracts.Where(v => v.Symbol == _Contract.Symbol && v.Expiry.Month == _Contract.Expiry.Month && v.Expiry.Year == _Contract.Expiry.Year).FirstOrDefault();
                            if (temp != null)
                                UnderlyingToken = temp.Token;
                            else
                            {
                                //var twmp = list_EQSecurity.Where(v => v.Symbol == _Contract.Symbol).FirstOrDefault();
                                //if (twmp != null)
                                //{
                                //    UnderlyingToken = twmp.Token;
                                //    USegment = "NSECM";
                                //}

                                var temp1 = FutContracts.Where(v => v.Symbol == _Contract.Symbol).OrderByDescending(v => v.Expiry).FirstOrDefault();

                                if (temp1 != null)
                                    UnderlyingToken = temp1.Token;

                            }
                        }

                        var pricefrombhavcopy = Bhavcopy.Where(v => v.CustomScripname == _Contract.CustomScripname).FirstOrDefault();//Added by Musharraf to fetch closing and settling price
                        if (pricefrombhavcopy != null)
                        {
                            ClosingPrice = pricefrombhavcopy.Close; //if null value is entered then Default should be zero
                            SettlingPrice = pricefrombhavcopy.SettlePrice;
                            OpenInterest = pricefrombhavcopy.OpenInterest;
                        }
                        //var ClosingPrice = (pricefrombhavcopy != null) ? pricefrombhavcopy.Close : 0; //if null value is entered then Default should be zero
                        //var SettlingPrice = (pricefrombhavcopy != null) ? pricefrombhavcopy.SettlePrice : 0;
                        //var OpenInterest = (pricefrombhavcopy != null) ? pricefrombhavcopy.OpenInterest : 0;
                        if (UnderlyingToken != -1)
                            list_ContractMasterRows.Add($"({_Contract.Token},'{_Contract.Symbol}','{_Contract.Instrument}','-','NSEFO','{_Contract.ScripName}'," +
                                $"'{_Contract.CustomScripname}','{_Contract.ScripType}','{_Contract.ExpiryUnix}',{_Contract.StrikePrice},{_Contract.LotSize}," +
                                $"{UnderlyingToken},'{USegment}',{ClosingPrice},{SettlingPrice},{OpenInterest})");//pricefrombhavcopy.Close, pricefrombhavcopy.SettlePrice
                    }

                    sb_InsertCommand.Append(string.Join(",", list_ContractMasterRows));

                    using (MySqlConnection myConnToken = new MySqlConnection(_MySQLCon))
                    {
                        using (MySqlCommand myCmd = new MySqlCommand(sb_InsertCommand.ToString(), myConnToken))
                        {
                            myConnToken.Open();
                            myCmd.CommandType = CommandType.Text;
                            myCmd.ExecuteNonQuery();
                            myConnToken.Close();
                        }
                    }

                    list_ComponentStarted.Add("FOUploadToken");
                    //FOTokenUploaded = true;
                }
                else
                {
                    AddToList("Contract file not found", true);
                    //FOTokenUploaded = false;
                }
            }
            catch (Exception ee)
            {
                _logger.Error(ee, " : InsertTokensIntoDB FO");

                AddToList("FO Tokens Upload failed. Please check logs for more details.", true);

                IsWorking = false;
                btn_RestartAuto.Enabled = true;
                btn_Settings.Enabled = true;
                //FOTokenUploaded = false;

                if (list_ComponentStarted.Contains("CMUploadToken"))
                    list_ComponentStarted.Remove("CMUploadToken");
                if (list_ComponentStarted.Contains("FOUploadToken"))
                    list_ComponentStarted.Remove("FOUploadToken");


                SentMail("FO Tokens");
            }

            #endregion

        }


        /// <summary>
        /// Reads the CD bhavcopy.
        /// </summary>
        /// <param name="CDBhavcopyPath">The CD bhavcopy path.</param>
        /// <param name="IsMar23Format">If true, is mar23 format.</param></param>
        /// <returns><![CDATA[List<CDBhavcopy>]]></returns>



        // Added by Snehadri on 15JUN2021 for Automatic BOD Process
        private void ReadContractMaster()
        {
            try
            {
                var dt_ContractMaster = new ds_Engine.dt_ContractMasterDataTable();
                using (MySqlConnection _mySqlConnection = new MySqlConnection(_MySQLCon))
                {
                    _mySqlConnection.Open();

                    using (MySqlCommand myCmdEod = new MySqlCommand("sp_GetContractMaster", _mySqlConnection))
                    {
                        myCmdEod.CommandType = CommandType.StoredProcedure;
                        using (MySqlDataAdapter dadapter = new MySqlDataAdapter(myCmdEod))
                        {
                            //changed on 13JAN2021 by Amey
                            dadapter.Fill(dt_ContractMaster);

                            dadapter.Dispose();
                        }
                    }

                    _mySqlConnection.Close();
                }

                ContractMaster ScripInfo = null;
                foreach (ds_Engine.dt_ContractMasterRow v in dt_ContractMaster.Rows)
                {

                    ScripInfo = new ContractMaster()
                    {
                        Token = v.Token,
                        Series = v.Series,
                        Symbol = v.Symbol,
                        InstrumentName = v.InstrumentName == "EQ" ? en_InstrumentName.EQ : (v.InstrumentName == "FUTIDX" ? en_InstrumentName.FUTIDX :
                        (v.InstrumentName == "FUTSTK" ? en_InstrumentName.FUTSTK : (v.InstrumentName == "OPTIDX" ? en_InstrumentName.OPTIDX : en_InstrumentName.OPTSTK))),
                        Segment = v.Segment == "NSECM" ? en_Segment.NSECM : (v.Segment == "NSECD" ? en_Segment.NSECD :
                        (v.Segment == "NSEFO" ? en_Segment.NSEFO : en_Segment.BSECM)),
                        ScripName = v.ScripName,
                        CustomScripName = v.CustomScripName,
                        ScripType = (v.ScripType == "EQ" ? n.Structs.en_ScripType.EQ : (v.ScripType == "XX" ? n.Structs.en_ScripType.XX : (v.ScripType == "CE" ? n.Structs.en_ScripType.CE :
                                    n.Structs.en_ScripType.PE))),
                        ExpiryUnix = v.ExpiryUnix,
                        StrikePrice = v.StrikePrice,
                        LotSize = v.LotSize,
                        UnderlyingToken = v.UnderlyingToken,
                        UnderlyingSegment = v.UnderlyingSegment == "NSECM" ? en_Segment.NSECM : (v.UnderlyingSegment == "NSECD" ? en_Segment.NSECD :
                        (v.UnderlyingSegment == "NSEFO" ? en_Segment.NSEFO : en_Segment.BSECM)),
                         //Snehadri - New - Primes
                        ClosingPrice = v.ClosingPrice,
                        SettlementPrice = v.SettlementPrice
                    };

                    dict_ScripInfo.TryAdd($"{ScripInfo.Segment}|{ScripInfo.ScripName}", ScripInfo);
                    dict_CustomScripInfo.TryAdd($"{ScripInfo.Segment}|{ScripInfo.CustomScripName}", ScripInfo);
                    dict_TokenScripInfo.TryAdd($"{ScripInfo.Segment}|{ScripInfo.Token}", ScripInfo);


                }

                list_ComponentStarted.Add("ReadContractMaster");
            }
            catch (Exception ee) { _logger.Error(ee, "ReadContractMaster : "); }
        }

        private void UploadClientMaster(string uploadtype, string UploadPath)
        {
            try
            {
                StringBuilder sCommand = new StringBuilder("INSERT INTO tbl_clientdetail(ClientID,DealerID,UserID,Username,Name,Margin,Adhoc,Zone,Branch,Family,Product) values(");
                int _rowsAffected = 0; DateTime dt_StartTime = DateTime.Now;//02-01-2020
                using (FileStream stream = File.Open(UploadPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    using (StreamReader reader = new StreamReader(stream))
                    {
                        string line1;
                        StringBuilder detail = new StringBuilder();
                        string[] uploadClients;
                        List<StringBuilder> Rows = new List<StringBuilder>();
                        while ((line1 = reader.ReadLine()) != null)
                        {
                            uploadClients = line1.Split(',');
                            if (uploadClients.Length >= 11)   //17-11-17
                            {
                                //Changed by Akshay on 08 - 01 - 2021 For new Columns
                                if ((uploadClients[0].Trim() == "" && uploadClients[1].Trim() == "" && uploadClients[2].Trim() == "") || uploadClients[3].Trim() == "" || uploadClients[4].Trim() == "" || uploadClients[5].Trim() == "" || uploadClients[6].Trim() == "" || uploadClients[7].Trim() == "" || uploadClients[8].Trim() == "" || uploadClients[9].Trim() == "" || uploadClients[10].Trim() == "")
                                {
                                    //added logs on 03MAY2021 by Amey
                                    _logger.Error(null, "Client Upload File Empty : " + line1);
                                    AddToList("Client Upload File Empty : " + line1, true);
                                    IsWorking = false;
                                    btn_RestartAuto.Enabled = true;
                                    btn_Settings.Enabled = true;
                                    SentMail("Client File upload failed");
                                    return;
                                }
                                _rowsAffected++;//02-01-2020

                                detail.Append("'" + uploadClients[0].ToUpper().Trim() + "',");   //ClientID
                                detail.Append("'" + uploadClients[1].ToUpper().Trim() + "',");   //DealerID
                                detail.Append("'" + uploadClients[2].ToUpper().Trim() + "',");   //UserID
                                detail.Append("'" + uploadClients[3].ToUpper().Trim() + "',");   //UserName
                                detail.Append("'" + uploadClients[4].ToUpper().Trim() + "',");   //Name

                                //added on 31DEC2020 by Amey
                                try
                                {
                                    decimal Margin = Convert.ToDecimal(uploadClients[5]);
                                    decimal AdHoc = Convert.ToDecimal(uploadClients[6]);

                                    if (AdHoc < 0)
                                    {
                                        _logger.Error(null, "Client Upload Loop Negetive : " + line1);
                                        AddToList("Please Check Client information in the Client File", true);
                                        IsWorking = false;
                                        btn_RestartAuto.Enabled = true;
                                        btn_Settings.Enabled = true;
                                        SentMail("Client File upload failed");
                                        return;
                                    }

                                    detail.Append("" + Margin + ",");
                                    detail.Append("" + AdHoc + ",");
                                }
                                catch (Exception ee)
                                {
                                    _logger.Error(ee, Environment.NewLine + "Client Upload Loop : " + line1);
                                    AddToList("Please Check Client information in the Client Data File", true);
                                    IsWorking = false;
                                    btn_Settings.Enabled = true;
                                    btn_RestartAuto.Enabled = true;
                                    SentMail("Client File upload failed");
                                    return;
                                }

                                detail.Append("'" + uploadClients[7].ToUpper().Trim() + "',");          //Zone          
                                detail.Append("'" + uploadClients[8].ToUpper().Trim() + "',");          //Branch        
                                detail.Append("'" + uploadClients[9].ToUpper().Trim() + "',");          //Family        
                                detail.Append("'" + uploadClients[10].ToUpper().Trim() + "'");           //Product 

                                detail.Append("),(");
                            }
                            else
                            {
                                //added logs on 03MAY2021 by Amey
                                _logger.Error(null, "Client Upload Loop Incomplete : " + line1);
                                AddToList("Incomplete client data", true);
                                IsWorking = false;
                                btn_Settings.Enabled = true;
                                btn_RestartAuto.Enabled = true;
                                SentMail("Client File upload failed");
                                return;
                            }
                        }
                        try
                        {
                            if (uploadtype == "Complete")
                            {
                                AddToList("Full Client upload started");

                                //added on 29JAN2021 by Amey
                                using (var con_MySQL = new MySqlConnection(_MySQLCon))
                                {
                                    con_MySQL.Open();
                                    using (var cmd = new MySqlCommand("TRUNCATE tbl_clientdetail ", con_MySQL))
                                    {
                                        cmd.ExecuteNonQuery();
                                        cmd.Dispose();
                                    }
                                }
                            }
                            else
                                AddToList("Partial Client upload started");
                        }
                        catch (Exception trunEx)
                        {
                            _logger.Error(null, "Upload client " + trunEx.ToString());
                            AddToList("Full Client upload failed. Please check the log", true);
                            IsWorking = false;
                            btn_RestartAuto.Enabled = true;
                            btn_Settings.Enabled = true;
                            SentMail("Client File upload failed");
                        }
                        //}
                        //added on 22-12-17 to add only proper rows with proper data format
                        Rows.Add(detail);
                        if (Rows.Count != 0)
                        {
                            sCommand.Append(string.Join(",", Rows));
                            int a = sCommand.Length;
                            sCommand.Remove(a - 2, 2);
                            sCommand.Append("ON DUPLICATE KEY UPDATE `Username`= VALUES(`Username`),`Name`= VALUES(`Name`),`Margin`= VALUES(`Margin`),`Adhoc`= VALUES(`Adhoc`),`Zone`= VALUES(`Zone`),`Branch`= VALUES(`Branch`),`Family`= VALUES(`Family`),`Product`= VALUES(`Product`);");
                            //sCommand.Append(";");

                            using (var con_MySQL = new MySqlConnection(_MySQLCon))
                            {
                                con_MySQL.Open();

                                using (MySqlCommand myCmd = new MySqlCommand(sCommand.ToString(), con_MySQL))
                                {
                                    myCmd.CommandType = CommandType.Text;
                                    myCmd.ExecuteNonQuery();
                                    myCmd.Dispose();
                                }
                            }
                        }
                    }
                }

                AddToList("Client File Uploaded Successfully. Row count " + _rowsAffected);
                _logger.Error(null, "Row count " + _rowsAffected + ", Total time taken for client upload " + (DateTime.Now - dt_StartTime));
                list_ComponentStarted.Add("ClientFileUpload");

            }
            catch (Exception ee)
            {
                _logger.Error(ee, "UploadClientMaster: ");
                IsWorking = false;
                btn_RestartAuto.Enabled = true;
                btn_Settings.Enabled = true;
                SentMail("Client File upload failed");
            }
        }

        //Added by Snehadri on 
        private void AddUserandClientMapping(string UserFile)
        {
            try
            {
                DataTable dt_UserInfo = new DataTable();
                using (var con_MySQL = new MySqlConnection(_MySQLCon))
                {
                    con_MySQL.Open();
                    using (MySqlCommand myCmd = new MySqlCommand("sp_GetUserInfo", con_MySQL))//modified by Navin on 12-06-2019
                    {
                        myCmd.CommandType = CommandType.StoredProcedure;

                        MySqlDataAdapter dadapt = new MySqlDataAdapter(myCmd);
                        dadapt.Fill(dt_UserInfo);
                    }
                }
                clsEncryptionDecryption.DecryptData(dt_UserInfo);


                var arr_Info = File.ReadAllLines(UserFile);

                foreach (var Info in arr_Info)
                {
                    if (Info != "")
                    {
                        var list_data = Info.Split(',').ToList();
                        var username = list_data[0].Trim();
                        var command = list_data[1].Trim().ToLower();

                        if (username == "")
                        {
                            AddToList("Invalid data in client mapping file", true);
                            IsWorking = false;
                            btn_RestartAuto.Enabled = true;
                            return;
                        }

                        List<string> list_MappedClients = new List<string>();

                        DataColumn[] columns = dt_UserInfo.Columns.Cast<DataColumn>().ToArray();
                        bool userpresent = dt_UserInfo.AsEnumerable().Any(row => columns.Any(col => row[col].ToString() == username));

                        if (userpresent)
                        {
                            DataRow user_row = dt_UserInfo.AsEnumerable().Where(row => row.Field<string>("UserName") == username).First();

                            if (command == "partial")
                            {
                                list_MappedClients = user_row[1].ToString().Split(',').ToList();
                                for (int i = 2; i < list_data.Count; i++)
                                {
                                    if (!list_MappedClients.Contains(list_data[i]))
                                        list_MappedClients.Add(list_data[i]);
                                }

                            }
                            else if (command == "delete")
                            {

                                list_MappedClients = user_row[1].ToString().Split(',').ToList();
                                for (int i = 2; i < list_data.Count; i++)
                                {
                                    if (list_MappedClients.Contains(list_data[i]))
                                        list_MappedClients.Remove(list_data[i]);
                                }

                            }
                            else if (command == "full")
                            {
                                for (int i = 2; i < list_data.Count; i++)
                                {
                                    if (!list_MappedClients.Contains(list_data[i]))
                                        list_MappedClients.Add(list_data[i]);
                                }
                            }
                            else if (command == "all")
                            {
                                using (var con_MySQL = new MySqlConnection(_MySQLCon))
                                {
                                    con_MySQL.Open();
                                    using (MySqlCommand myCmd = new MySqlCommand("sp_GetClientDetail", con_MySQL))
                                    {
                                        myCmd.CommandType = CommandType.StoredProcedure;

                                        myCmd.Parameters.Add("prm_Type", MySqlDbType.LongText);
                                        myCmd.Parameters["prm_Type"].Value = "ALL";

                                        using (MySqlDataReader mySqlDataReader = myCmd.ExecuteReader())
                                        {
                                            while (mySqlDataReader.Read())
                                            {
                                                list_MappedClients.Add(mySqlDataReader.GetString(3));

                                            }
                                        }
                                    }
                                }
                            }
                            else
                            {
                                AddToList("Invalid data in client mapping file", true);
                                IsWorking = false;
                                btn_RestartAuto.Enabled = true;
                                return;
                            }


                            StringBuilder sbMappedClient = new StringBuilder();
                            if (list_MappedClients.Count > 0 || command == "delete")
                            {
                                for (int i = 0; i < list_MappedClients.Count; i++)
                                {
                                    sbMappedClient.Append(list_MappedClients[i] + ",");
                                }
                                if (sbMappedClient.Length > 1)
                                    sbMappedClient.Remove(sbMappedClient.ToString().LastIndexOf(','), 1);

                                using (var con_MySQL = new MySqlConnection(_MySQLCon))
                                {
                                    con_MySQL.Open();
                                    using (var cmd = new MySqlCommand("UPDATE tbl_login SET MappedClient = '" + clsEncryptionDecryption.EncryptString(sbMappedClient.ToString(), "Nerve123") + "' where UserName= '" + clsEncryptionDecryption.EncryptString(username.ToLower(), "Nerve123") + "'", con_MySQL))
                                    {
                                        int result = cmd.ExecuteNonQuery();
                                        if (result != 1)
                                        {
                                            AddToList("Client Mapping failed for User: " + username, true);
                                            IsWorking = false;
                                            btn_RestartAuto.Enabled = true;
                                            if (list_ComponentStarted.Contains("Clientmapping"))
                                                list_ComponentStarted.Remove("Clientmapping");
                                            return;
                                        }

                                    }
                                }
                            }



                        }
                        else
                        {

                            if (command == "all")
                            {
                                using (var con_MySQL = new MySqlConnection(_MySQLCon))
                                {
                                    con_MySQL.Open();
                                    using (MySqlCommand myCmd = new MySqlCommand("sp_GetClientDetail", con_MySQL))
                                    {
                                        myCmd.CommandType = CommandType.StoredProcedure;

                                        myCmd.Parameters.Add("prm_Type", MySqlDbType.LongText);
                                        myCmd.Parameters["prm_Type"].Value = "ALL";

                                        using (MySqlDataReader mySqlDataReader = myCmd.ExecuteReader())
                                        {
                                            while (mySqlDataReader.Read())
                                            {
                                                list_MappedClients.Add(mySqlDataReader.GetString(3));

                                            }
                                        }
                                    }
                                }

                            }
                            else
                            {
                                for (int i = 2; i < list_data.Count; i++)
                                {
                                    list_MappedClients.Add(list_data[i]);
                                }

                            }

                            StringBuilder sbMappedClient = new StringBuilder();
                            if (list_MappedClients.Count > 0)
                            {

                                for (int i = 0; i < list_MappedClients.Count; i++)
                                {
                                    sbMappedClient.Append(list_MappedClients[i] + ",");
                                }
                                if (sbMappedClient.Length > 1)
                                    sbMappedClient.Remove(sbMappedClient.ToString().LastIndexOf(','), 1);

                                int result = 0;
                                using (var con_MySQL = new MySqlConnection(_MySQLCon))
                                {
                                    con_MySQL.Open();
                                    using (var cmd = new MySqlCommand("INSERT INTO `tbl_login` (`UserName`, `Password`, `IsAdmin`, MappedClient) VALUES ( '" + clsEncryptionDecryption.EncryptString(username.ToLower(), "Nerve123") + "','" + clsEncryptionDecryption.EncryptString("prime123", "Nerve123") + "','" + clsEncryptionDecryption.EncryptString("false", "Nerve123") + "','" + clsEncryptionDecryption.EncryptString(sbMappedClient.ToString(), "Nerve123") + "')", con_MySQL))
                                    {
                                        result = cmd.ExecuteNonQuery();
                                        if (result != 1)
                                        {
                                            AddToList("Client Mapping failed for User: " + username, true);
                                            IsWorking = false;
                                            btn_RestartAuto.Enabled = true;
                                            if (list_ComponentStarted.Contains("Clientmapping"))
                                                list_ComponentStarted.Remove("Clientmapping");
                                            return;
                                        }
                                    }
                                }
                            }


                        }
                    }

                }

                AddToList("Mapping Clients completed successfully");
                list_ComponentStarted.Add("Clientmapping");
            }
            catch (Exception ee)
            {
                _logger.Error(ee, "AddUserandClientMapping: ");
                IsWorking = false;
                btn_RestartAuto.Enabled = true;
                SentMail("Client Mapping failed to execute");
                AddToList("Mapping Clients unsuccessful", true);
                if (list_ComponentStarted.Contains("Clientmapping"))
                    list_ComponentStarted.Remove("Clientmapping");
            }
        }



        // Added by Snehadri on 15JUN2021 for Automatic BOD Process
        private void ClearEOD()
        {
            try
            {
                AddToList("Deleting EOD data");

                Application.DoEvents();
                using (MySqlConnection myConnClear = new MySqlConnection(_MySQLCon))
                {
                    using (MySqlCommand myCmdClear = new MySqlCommand("sp_ClearEOD", myConnClear))
                    {
                        myConnClear.Open();
                        myCmdClear.CommandType = CommandType.StoredProcedure;
                        myCmdClear.ExecuteNonQuery();
                        myConnClear.Close();
                    }
                }

                AddToList("EOD data deleted successfully.");
                list_ComponentStarted.Add("ClearEOD");

            }
            catch (Exception ee)
            {
                _logger.Error(ee);
                AddToList("EOD data deletion unsuccessfull. Please check logs for more details.", true);
                IsWorking = false;
                btn_RestartAuto.Enabled = true;
                btn_Settings.Enabled = true;
                SentMail("Clear EOD");
            }
        }

        // Added by Snehadri on 15JUN2021 for Automatic BOD Process
        private void SelectClientfromDatabase()
        {
            try
            {
                using (MySqlConnection con_MySQL = new MySqlConnection(_MySQLCon))
                {
                    using (MySqlCommand myCmd = new MySqlCommand("sp_GetClientDetail", con_MySQL))
                    {
                        myCmd.CommandType = CommandType.StoredProcedure;

                        //added on 27APR2021 by Amey
                        myCmd.Parameters.Add("prm_Type", MySqlDbType.LongText);
                        myCmd.Parameters["prm_Type"].Value = "ID";

                        con_MySQL.Open();

                        using (MySqlDataReader _mySqlDataReader = myCmd.ExecuteReader())
                        {
                            while (_mySqlDataReader.Read())
                            {
                                string ClientID = _mySqlDataReader.GetString(0).ToUpper().Trim();

                                //added on 12JAN2021 by Amey
                                hs_Usernames.Add(ClientID);
                            }
                        }

                        con_MySQL.Close();
                    }
                }
            }
            catch (Exception clientEx)
            {
                _logger.Error(null, "SelectClientfromDatabase " + clientEx);
            }
        }

        // Added by Snehadri on 15JUN2021 for Automatic BOD Process
        private void InsertDay1(string filename)
        {
            try
            {
                //added Segment on 20APR2021 by Amey. To avoid same Token conflict from different segments.
                var result = list_Day1Positions.GroupBy(s => new { s.Token, s.Username, s.Segment })
                                            .Select(g => new
                                            {
                                                Username = g.Select(x => x.Username).First(),
                                                Segment = g.Select(x => x.Segment).First(),
                                                Token = g.Select(x => x.Token).First(),
                                                BEP = Math.Round(g.Sum(x => x.TradeQuantity * x.TradePrice) / (g.Sum(x => x.TradeQuantity) == 0 ? -1 : g.Sum(x => x.TradeQuantity)), 4),
                                                TradeQuantity = g.Sum(x => x.TradeQuantity),
                                                UnderlyingToken = g.Select(x => x.UnderlyingToken).First(),
                                                UnderlyingSegment = g.Select(x => x.UnderlyingSegment).First()
                                            });

                MySqlCommand cmd = new MySqlCommand();

                //changed on 07JAN2021 by Amey
                StringBuilder insertCmd = new StringBuilder("INSERT IGNORE INTO tbl_eod (Username,Segment,Token,TradePrice,TradeQuantity,UnderlyingSegment,UnderlyingToken) VALUES");

                List<string> toInsert = new List<string>();

                //changed to var on 27APR2021 by Amey
                var date_Tick = ConvertToUnixTimestamp(DateTime.Now);

                foreach (var _Item in result)
                {
                    var ScripInfo = dict_TokenScripInfo[$"{_Item.Segment}|{_Item.Token}"];
                    if ((ScripInfo.ExpiryUnix) > date_Tick || ScripInfo.ScripType == n.Structs.en_ScripType.EQ)   //09-01-18
                        toInsert.Add($"('{_Item.Username}','{ScripInfo.Segment}',{ScripInfo.Token},{_Item.BEP},{_Item.TradeQuantity},'{ScripInfo.UnderlyingSegment}',{ScripInfo.UnderlyingToken})");
                }
                try
                {
                    if (toInsert.Count > 0)
                    {
                        insertCmd.Append(string.Join(",", toInsert));
                        insertCmd.Append(";");
                        using (MySqlConnection myconnDay1 = new MySqlConnection(_MySQLCon))
                        {
                            cmd = new MySqlCommand(insertCmd.ToString(), myconnDay1);
                            myconnDay1.Open();
                            cmd.ExecuteNonQuery();
                            myconnDay1.Close();
                        }
                    }

                    AddToList($"{filename} Process Completed. {toInsert.Count} Rows added.");
                    _logger.Debug($"{filename} Process Completed. {toInsert.Count} Rows added.");

                    list_Day1Positions.Clear();
                }
                catch (Exception ee)
                {
                    _logger.Error(ee, "InsertDay1 -inner");
                    AddToList($"{filename} Process failed. Please check the log file.", true);
                }
            }
            catch (Exception ee)
            {
                _logger.Error(ee, "InsertDay1");
                AddToList($"{filename} Process failed. Please check the log file.", true);
            }
        }

        // Added by Snehadri on 15JUN2021 for Automatic BOD Process
        private void Day1andPS03FileUpload(string DAY1Folder, string PS03Folder, string BhavcopyPath)
        {
            try
            {
                bool Day1_error = false; bool PS03_error = false; /*bool PS04_error = false;*/

                try
                {
                    SelectClientfromDatabase();

                    var Day1Directory = new DirectoryInfo(DAY1Folder);

                    var Day1File = Day1Directory.GetFiles()
                                                  .OrderByDescending(f => f.LastWriteTime)
                                                  .First();
                    AddToList("Day1 Process started.");

                    //Seperated class for reading Day1 Positions for better track code updates of various Prime versions. 09MAR2021-Amey
                    list_Day1Positions = Day1.Read(DAY1Folder, BhavcopyPath, _logger, true, true, hs_Usernames, dict_ScripInfo, dict_CustomScripInfo, dict_TokenScripInfo);


                    if (Day1.Filename != "")
                        AddToList($"Day1 Filename : {Day1.Filename}");

                    //addded on 09APR2021 by Amey
                    if (Day1.isAnyError)
                    {
                        AddToList("Invalid data found in Day1 file. Please check if positions are uploaded properly.", true);
                        Day1_error = true;

                    }
                    else
                    {
                        InsertDay1("Day1");
                        Day1_error = false;
                    }
                }
                catch (Exception ee)
                {
                    _logger.Error(ee, "Day1 File Upload Error.");
                    Day1_error = true;
                }


                try
                {
                    var PS03directory = new DirectoryInfo(PS03Folder);
                    var PS03 = PS03directory.GetFiles().OrderByDescending(f => f.LastWriteTime).First();
                    AddToList("PS03 Process started.");

                    //Seperated class for reading Day1 Positions for better track code updates of various Prime versions. 09MAR2021-Amey
                    list_Day1Positions = Day1.ReadPS03(PS03Folder, BhavcopyPath, _logger, true, true, hs_Usernames, dict_ScripInfo, dict_CustomScripInfo, dict_TokenScripInfo);

                    if (Day1.Filename != "")
                        AddToList($"PS03 Filename : {Day1.Filename}");

                    //addded on 09APR2021 by Amey
                    if (Day1.isAnyError)
                    {
                        AddToList("Invalid data found in PS03 file. Please check if positions are uploaded properly.", true);
                        PS03_error = true;

                    }
                    else
                    {
                        InsertDay1("PS03");
                        PS03_error = false;
                    }
                }
                catch (Exception ee)
                {
                    _logger.Error(ee, "PS03 File Upload Error.");
                    PS03_error = true;
                }

                if (!Day1_error || !PS03_error)
                {
                    IsWorking = true;
                    list_ComponentStarted.Add("Day1PS03File");
                }
                else
                {
                    IsWorking = false;
                    btn_RestartAuto.Enabled = true;
                    btn_Settings.Enabled = true;
                    SentMail("Day1/PS03 file upload error.");
                }

            }
            catch (Exception ee)
            {
                _logger.Error(ee);
                AddToList("Error in Uploading Day1/PS03 file. Please check logs for more details.", true);
                IsWorking = false;
                btn_RestartAuto.Enabled = true;
                btn_Settings.Enabled = true;
                SentMail("Day1 or PS03 file failed to upload.");
            }

        }

        // Added by Snehadri on 15JUN2021 for Automatic BOD Process
        private void StartGateway(string GatewayPath)
        {
            try
            {
                CloseComponentexe("Gateway");

                OpenComponentexe("Gateway", GatewayPath);

                bool gatewaystarted = GatewayEngineConnector.ConnectComponents("Gateway");
                if (gatewaystarted) { AddToList("Gateway Started"); list_ComponentStarted.Add("Gateway"); }
                else
                {
                    AddToList("Gateway has failed to launch, Please check the log", true);
                    IsWorking = false;
                    btn_RestartAuto.Enabled = true;
                    btn_Settings.Enabled = true;
                    SentMail("Gateway has failed to launch");
                }
            }
            catch (Exception ee)
            {
                _logger.Error(ee);
                AddToList("Gateway has failed to launch, Please check the log", true);
                IsWorking = false;
                btn_RestartAuto.Enabled = true;
                btn_Settings.Enabled = true;
                SentMail("Gateway has failed to launch");
            }

        }

        // Added by Snehadri on 15JUN2021 for Automatic BOD Process
        private void StartEngine(string EnginePath)
        {
            try
            {
                CloseComponentexe("Engine");

                OpenComponentexe("Engine", EnginePath);

                Thread.Sleep(5000);         // This sleep is to allow the n.Engine to initialise 
                bool gatewaystarted = GatewayEngineConnector.ConnectComponents("Engine");
                if (gatewaystarted) { AddToList("Engine Started"); list_ComponentStarted.Add("Engine"); }
                else
                {
                    AddToList("Engine Not Started, Please check the log", true);
                    IsWorking = false;
                    btn_RestartAuto.Enabled = true;
                    btn_Settings.Enabled = true;
                    SentMail("Engine has failed to launch");
                }
            }
            catch (Exception ee)
            {
                _logger.Error(ee);
                AddToList("Engine has failed to launch, Please check the log", true);
                IsWorking = false;
                btn_RestartAuto.Enabled = true;
                btn_Settings.Enabled = true;
                SentMail("Engine has failed to launch");
            }
        }

        public static DateTime ConvertFromUnixTimestamp(double timestamp)
        {
            DateTime origin = new DateTime(1980, 1, 1, 0, 0, 0, 0, DateTimeKind.Utc);
            return origin.AddSeconds(timestamp);
        }

        // Added by Snehadri on 15JUN2021 for Automatic BOD Process
        private void btn_Settings_Click(object sender, EventArgs e)
        {
            try
            {
                new Settings().ShowDialog();
            }
            catch (Exception ee) { _logger.Error(ee); }
        }

        private void OpenComponentexe(string componentname, string filepath)
        {
            try
            {
                Process process = new Process();
                process.StartInfo.FileName = filepath;
                process.StartInfo.UseShellExecute = false;
                process.StartInfo.WorkingDirectory = Path.GetDirectoryName(filepath);
                process.Start();
            }
            catch (Exception ee) { _logger.Error(ee, $"{componentname} has failed to start."); }
        }

        private void CloseComponentexe(string componentname)
        {
            try
            {
                string[] array = componentname.Split(',');
                foreach (var item in array)
                {
                    Process[] process = Process.GetProcessesByName(item.Trim());
                    if (process.Length > 0)
                    {
                        foreach (var prog in process)
                        {
                            prog.Kill();
                        }
                        Thread.Sleep(5000);
                    }

                }


            }
            catch (Exception ee) { _logger.Error(ee); }
        }

        #region PS03
        //private void ConvertPSO3Files()
        //{
        //    try
        //    {
        //        var dRow = ds_Config.Tables["AUTOMATICSETTINGS"].Rows[0];
        //        var NewPS03FolderPath = dRow["PS0-FOLDER-NEW-FORMAT"].ToString();

        //        var NewPS03directory = new DirectoryInfo(NewPS03FolderPath);
        //        var NewPS03_Files = NewPS03directory.GetFiles().OrderByDescending(f => f.LastWriteTime);

        //        foreach (var NewPS03 in NewPS03_Files)
        //        {
        //            var arr_lines = File.ReadAllLines(NewPS03.FullName);

        //            StringBuilder sb_OldFile = new StringBuilder();
        //            foreach (var line in arr_lines)
        //            {
        //                if (line.Contains("ClntId"))
        //                {
        //                    continue;
        //                }
        //                var arr_values = line.Split(',');
        //                for (int i = 0; i < arr_values.Length; i++)
        //                {
        //                    if (i == 10 || i == 15 || i == 16)
        //                        continue;
        //                    sb_OldFile.Append(arr_values[i] + ",");
        //                }
        //                sb_OldFile.Append('\n');
        //            }

        //            if (NewPS03.Name.Contains("PS04"))
        //            {
        //                try
        //                {
        //                    var OldFolderPath = dRow["PSO4FOLDER"].ToString();
        //                    if (!OldFolderPath.EndsWith("\\"))
        //                    {
        //                        OldFolderPath += "\\";
        //                    }
        //                    var OldPso3File = OldFolderPath + "F_PS04_" + DateTime.Now.ToString("ddMMyyyy") + ".csv";
        //                    if (NewPS03.Name.Contains("BFX"))
        //                    {
        //                        OldPso3File = OldFolderPath + "X_PS04_" + DateTime.Now.ToString("ddMMyyyy") + ".csv";
        //                    }
        //                    File.WriteAllText(OldPso3File, sb_OldFile.ToString());
        //                }
        //                catch (Exception ee) { _logger.Error(ee); }

        //            }
        //            else if (NewPS03.Name.Contains("PS03"))
        //            {
        //                try
        //                {
        //                    var OldFolderPath = dRow["PSO3FOLDER"].ToString();
        //                    if (!OldFolderPath.EndsWith("\\"))
        //                    {
        //                        OldFolderPath += "\\";
        //                    }
        //                    var OldPso3File = OldFolderPath + "F_PS03_" + DateTime.Now.ToString("ddMMyyyy") + ".csv";
        //                    if (NewPS03.Name.Contains("BFX"))
        //                    {
        //                        OldPso3File = OldFolderPath + "X_PS03_" + DateTime.Now.ToString("ddMMyyyy") + ".csv";
        //                    }
        //                    File.WriteAllText(OldPso3File, sb_OldFile.ToString());
        //                }
        //                catch (Exception EE)
        //                {
        //                    _logger.Error(EE);
        //                }

        //            }
        //        }
        //    }
        //    catch (Exception ee) { _logger.Error(ee); }
        //}
        #endregion



        // Added by Snehadri on 15JUN2021 for Automatic BOD Process
        private void StartComponents(bool IsRestart = false)            // Added by Snehadri on
        {
            try
            {
                // To avoid interference of the Engine's different socket connections 
                CloseComponentexe("Engine");

                if (!IsRestart)
                {
                    CloseComponentexe("CM FeedReceiver,FO FeedReceiver,CD FeedReceiver,NOTIS API EQ Manager,NOTIS API FO Manager,NOTIS API CD Manager,Gateway");

                    GatewayEngineConnector.StartServer();
                    Thread.Sleep(2000);
                }

                AddToList("Starting n.Prime Components");

                DataSet ds_SettingConfig = NerveUtils.XMLC(ApplicationPath + "config.xml");
                var dRow = ds_SettingConfig.Tables["AUTOMATICSETTINGS"].Rows[0];
                string[] CDFEEDPath = dRow["CDFEEDPATH"].STR().SPL(',');
                string[] CMFEEDPath = dRow["CMFEEDPATH"].STR().SPL(',');
                string[] FOFEEDPath = dRow["FOFEEDPATH"].STR().SPL(',');
                string[] NOTISEQPath = dRow["NOTISEQPATH"].STR().SPL(',');
                string[] NOTISFOPath = dRow["NOTISFOPATH"].STR().SPL(',');
                string[] NOTISCDPath = dRow["NOTISCDPATH"].STR().SPL(',');
                string DAY1Folder = dRow["DAY1FOLDER"].STR();
                string PS03Folder = dRow["PSO3FOLDER"].STR();
                string BhavcopyPath = dRow["BHAVCOPYPATH"].STR();
                string StartApi = dRow["STARTNOTIS"].STR().ToLower();
                string StartCDComponents = dRow["START_CD_COMPONENTS"].STR().ToLower();
                string GatewayPath = dRow["GATEWAYPATH"].STR();
                string EnginePath = dRow["ENGINEPATH"].STR();
                string ClientFullUploadPath = dRow["CLIENTFULLUPLOADPATH"].STR();
                string ClientPartialUploadPath = dRow["CLIENTPARTIALUPLOADPATH"].STR();
                string UserMappingFilePath = dRow["USERMAPPINGFILEPATH"].STR();

                if (IsWorking && !list_ComponentStarted.Contains("CMFeedReceiver"))
                {
                    StartCMFeedReceivers(CMFEEDPath);
                }

                if (IsWorking && !list_ComponentStarted.Contains("FOFeedReceiver"))
                {
                    StartFOFeedReceiver(FOFEEDPath);
                }

                if ((IsWorking && StartCDComponents == "yes") && !list_ComponentStarted.Contains("CDFeedReceiver"))
                {
                    StartCDFeedReceiver(CDFEEDPath);
                }

                if (StartApi == "yes")
                {
                    if (IsWorking && !list_ComponentStarted.Contains("NOTISEQReceiver"))
                    {
                        StartCMNotisApi(NOTISEQPath);
                    }

                    if (IsWorking && !list_ComponentStarted.Contains("NOTISFOReceiver"))
                    {
                        StartFONotisApi(NOTISFOPath);
                    }

                    if ((StartCDComponents == "yes" && IsWorking) && !list_ComponentStarted.Contains("NOTISCDReceiver"))
                    {
                        StartCDNotisApi(NOTISCDPath);
                    }
                }

                if (IsWorking && !list_ComponentStarted.Contains("CMUploadToken") && !list_ComponentStarted.Contains("FOUploadToken"))   // To start the Upload Token Procedure 
                {
                    InsertTokensIntoDBUdiff();
                    Thread.Sleep(5000);
                }

                if (IsWorking && !list_ComponentStarted.Contains("Gateway")) { StartGateway(GatewayPath); Thread.Sleep(5000); }     // To start Gateway

                if (IsWorking && !list_ComponentStarted.Contains("ClientFileUpload"))
                {
                    try
                    {
                        var FullUpload = new DirectoryInfo(ClientFullUploadPath);
                        var FullUploadFile = FullUpload.GetFiles()
                                   .OrderByDescending(f => f.LastWriteTime)
                                   .First();

                        UploadClientMaster("Complete", FullUploadFile.FullName);
                        Thread.Sleep(2000);

                    }
                    catch (Exception ee) { }
                    try
                    {
                        var PartialUpload = new DirectoryInfo(ClientPartialUploadPath);

                        var PartialUploadFile = PartialUpload.GetFiles()
                                   .OrderByDescending(f => f.LastWriteTime)
                                   .First();

                        UploadClientMaster("Partial", PartialUploadFile.FullName);
                        Thread.Sleep(2000);
                    }
                    catch (Exception ee) { }

                }

                if (IsWorking && !list_ComponentStarted.Contains("ClientMapped"))      // To Add User and Mapp the Clients 
                {
                    try
                    {
                        var UserMapping = new DirectoryInfo(UserMappingFilePath);
                        var UserMappingFile = UserMapping.GetFiles()
                                   .OrderByDescending(f => f.LastWriteTime)
                                   .First();
                        AddUserandClientMapping(UserMappingFile.FullName);
                    }
                    catch (Exception ee) { }
                }

                if (IsWorking && !list_ComponentStarted.Contains("ReadContractMaster")) { ReadContractMaster(); Thread.Sleep(5000); }     // Read the contract master from the database

                if (IsWorking && !list_ComponentStarted.Contains("ClearEOD")) { ClearEOD(); Thread.Sleep(5000); }                                  // To clear EOD

                if (IsWorking && !list_ComponentStarted.Contains("Day1PS03File")) // To Upload Day1 and PS03 File
                {
                    Day1andPS03FileUpload(DAY1Folder, PS03Folder, BhavcopyPath);
                    Thread.Sleep(5000);
                }

                if (IsWorking && !list_ComponentStarted.Contains("Engine")) { StartEngine(EnginePath); Thread.Sleep(5000); }                                                                               // To start Engine


                if (IsWorking)
                {
                    AddToList("BOD process successfull");
                    SentMail(null, false);
                    btn_Settings.Enabled = true;
                    GatewayEngineConnector.CloseServer();
                }
            }
            catch (Exception ee) { _logger.Error(ee); }
        }

        // Added by Snehadri on 15JUN2021 for Automatic BOD Process
        private void SentMail(string error_message, bool Isfault = true)
        {
            try
            {
                DataSet ds_SettingConfig = NerveUtils.XMLC(ApplicationPath + "config.xml");
                var dRow = ds_SettingConfig.Tables["AUTOMATICSETTINGS"].Rows[0];
                string sent_from = dRow["FROMEMAIL"].STR();
                string sent_to = dRow["TOEMAIL"].STR();
                string password = dRow["PASSWORD"].STR();
                string smtp = dRow["SMTP"].STR();
                string subject = "Automatic BOD process notifiation";
                string message = null;

                if (Isfault)
                {
                    message += $"Hi, \n There was a problem in starting the BOD process.\n" + error_message;
                }
                else
                {
                    message += "Hi, \n BOD process completed succefully";
                }

                //Sending Email
                SmtpClient client = new SmtpClient()
                {
                    Host = smtp,
                    Port = 587,
                    EnableSsl = true,
                    DeliveryMethod = SmtpDeliveryMethod.Network,
                    UseDefaultCredentials = false,
                    Credentials = new NetworkCredential()
                    {
                        UserName = sent_from,
                        Password = password,
                    }
                };
                MailAddress FromEmail = new MailAddress(sent_from);
                MailAddress ToEmail = new MailAddress(sent_to);
                MailMessage Message = new MailMessage()
                {
                    From = FromEmail,
                    Subject = subject,
                    Body = message,

                };
                Message.To.Add(ToEmail);

                try
                {
                    client.Send(Message);
                    AddToList("Email has been sent successfully");
                }
                catch (Exception ee) { _logger.Error(ee); AddToList("Please check the email configuration in the setting", true); }


            }
            catch (Exception ee) { _logger.Error(ee); }
        }

        private double ConvertToUnixTimestamp(DateTime date)
        {
            DateTime origin = new DateTime(1980, 1, 1, 0, 0, 0, 0, DateTimeKind.Utc);
            TimeSpan diff = date - origin;
            return diff.TotalSeconds;
        }

        // Added by Snehadri on 15JUN2021 for Automatic BOD Process
        private void btn_RestartAuto_Click(object sender, EventArgs e)
        {
            try
            {
                btn_RestartAuto.Enabled = false;
                IsWorking = true;
                AddToList("Automatic BOD Process Restarted");
                StartComponents(true);

            }
            catch (Exception ee) { _logger.Error(ee); }

        }

        private void btn_DownloadSpan_Click(object sender, EventArgs e)
        {
            try
            {
                _logger.Error(null, "Download Span Button Clicked");

                object tempObj = null;
                ElapsedEventArgs tempE = null;

                var arr_SpanInfo = ds_Config.GET("SAVEPATH", "SPAN").SPL(',');

                arr_SpanFileExtensions = ds_Config.GET("SAVEPATH", "SPAN-EXTENSTIONS").SPL(',');//Added by Musharraf to Download Span manually

                Task.Run(() => DownloadSpan(tempObj, tempE, arr_SpanInfo));

            }
            catch (Exception ee) { _logger.Error(ee); }
        }

        private void Home_Load(object sender, EventArgs e)
        {
            try
            {
                btn_DownloadSpan.Enabled = false;
                btn_RestartAuto.Enabled = false;
            }
            catch (Exception ee) { _logger.Error(ee); }
        }
    }
}
