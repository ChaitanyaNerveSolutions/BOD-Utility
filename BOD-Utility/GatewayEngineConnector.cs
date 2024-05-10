using NerveLog;
using NerveUtility;
using System;
using System.Data;
using System.Net;
using System.Net.Sockets;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using System.Xml;

namespace BOD_Utility                           // Added by Snehadri on 15JUN2021 for Automatic BOD Process
{
    public class GatewayEngineConnector
    {
        public static string ApplicationPath = Application.StartupPath + "\\";
        public static NerveLogger _logger = new NerveLogger(true, true, ApplicationName: "BOD-Utility");
        public static DataSet ds_Config = new DataSet();
        public static DataSet ds_MainConfig = new DataSet();


        public static IPAddress ipAddress;
        public static IPEndPoint localEndPoint;

        public static Socket BODServer;

        public static int config_Attempts = 0;
        public static int config_Interval = 0;

        public string IP;
        public int PORT;

        public void StartServer()
        {
            try
            {
                _logger.Initialize(ApplicationPath);
                ds_Config = NerveUtils.XMLC(ApplicationPath + "config.xml");
                var dRow = ds_Config.Tables["AUTOMATICSETTINGS"].Rows[0];
                config_Attempts += int.Parse(dRow["ATTEMPTS"].STR());
                config_Interval += int.Parse(dRow["INTERVAL"].STR()) * 1000;


                XmlTextReader tReader = new XmlTextReader("C:/Prime/Config.xml");
                tReader.Read();
                ds_MainConfig.ReadXml(tReader);
                var conn_info = ds_MainConfig.Tables["CONNECTION"].Rows[0];

                IP = conn_info["BOD-IP"].ToString();
                PORT = int.Parse(conn_info["BOD-PORT"].ToString());
                ipAddress = IPAddress.Parse(IP);
                localEndPoint = new IPEndPoint(ipAddress, PORT);
                BODServer = new Socket(ipAddress.AddressFamily, SocketType.Stream, ProtocolType.Tcp);

                BODServer.Bind(localEndPoint);
                BODServer.Listen(0);

            }
            catch (Exception error)
            {
                _logger.Error(error);
            }
        }

        public bool ConnectComponents(string componentname)
        {
            try
            {

                byte[] bytes = new byte[1024];
                string data = null;
                int attempt = 0;

                Socket server = new Socket(ipAddress.AddressFamily, SocketType.Stream, ProtocolType.Tcp);
                new Thread(() => { server = BODServer.Accept(); }).Start();
                data = null;
                while (attempt < config_Attempts)
                {
                    if (server.Connected)
                    {
                        int bytesRec = server.Receive(bytes);
                        data += Encoding.ASCII.GetString(bytes, 0, bytesRec);
                        if (data == $"{componentname}Started")
                        {
                            byte[] msg = Encoding.ASCII.GetBytes("START");
                            server.Send(msg);
                            server.Shutdown(SocketShutdown.Both);
                            break;
                        }
                        else
                        {
                            server.Shutdown(SocketShutdown.Both);
                            break;

                        }
                    }
                    attempt += 1;
                    Thread.Sleep(config_Interval);
                }
                server.Close();

                if (data == $"{componentname}Started")
                {
                    return true;
                }
                else
                {
                    _logger.Error(null, $"Error while connecting to {componentname}");
                    return false;
                }

            }
            catch (Exception error)
            {
                _logger.Error(error);
                return false;
            }
        }

        public void CloseServer()
        {
            try
            {
                BODServer.Close();
            }

            catch (Exception error)
            {
                _logger.Error(error);
            }
        }
    }


}