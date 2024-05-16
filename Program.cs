using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Threading.Tasks;
using System.Net.Http;
using System.Data;
using System.Data.OleDb;
using System.Security.Cryptography;
using System.Text.Json;
using System.Threading;
using System.Reflection;
using RestSharp;
using System.Diagnostics;
using static System.Net.Mime.MediaTypeNames;
using System.Web.UI.WebControls;
using System.Net.NetworkInformation;

namespace FingerScan
{
    public class PostData
    {
        public string USERID { get; set; }
        public string USERNAME { get; set; }
        public string CHECKTIME { get; set; }
    }
    public class Absensi
    {
        public PostData CheckInOut { get; set; }
    }

    internal class Program
    {
        static string connStringAbsensi = "";
        static string baseURL = "";
        static int maxRecords = 0;
        static int pauseEveryNSecs = 0;
        static string machineId = "X";
        static string logInfo = "";
        static string completeLogText = "";
        static volatile bool exit = false;

        static void Main(string[] args)
        {
            readConfigFile();

            logInfo = "Started at " + DateTime.Now;
            PrintNLog(logInfo);

            cekAbsensi();
        }

        static void readConfigFile()
        {
            StreamReader sr = new StreamReader("FingerScan.ini");
            string line = sr.ReadLine();

            while (line != null)
            {
                string kar = line.Substring(0, 1);
                if ("#'".Contains(kar) == false)
                {
                    var split = line.Split('=');
                    string keyName = split[0].Trim();
                    string value = split[1].Trim();

                    switch (keyName.ToLower())
                    {
                        //case "historydatabase":
                        //    connStringHistory = @"provider=Microsoft.Jet.OLEDB.4.0; Data Source=" + value;
                        //    break;
                        case "databaselocation":
                            connStringAbsensi = @"provider=Microsoft.Jet.OLEDB.4.0; Data Source=" + value;
                            break;
                        //case "machineid":
                        //    machineId = value.ToString();
                        //    break;
                        case "odoourl":
                            baseURL = value.ToString().Replace(@"\", "/");
                            break;
                        //case "maxrecords":
                        //    maxRecords = Convert.ToInt16(value);
                        //    break;
                        //case "pauseevery":
                        //    pauseEveryNSecs = Convert.ToInt16(value);
                        //    break;
                        default: break;
                    }

                    //logInfo = keyName + ": '" + value + "'";
                    //PrintNLog(logInfo);
                }

                line = sr.ReadLine();   //read the next line
            }
            sr.Close();
        }

        static bool cekKoneksiOdoo()
        {
            string info = "";
            try
            {
                string url = baseURL.Replace(@"http:", "");
                url = url.Substring(0, url.IndexOf(":")).Replace("//", "").Replace(@"\\", "");
                Ping myPing = new Ping();

                for (int i = 0; i <= 50; i++)
                {
                    PingReply reply = myPing.Send(url, 1000);
                    if (reply != null && reply.Status.ToString() == "Success")
                    {
                        info = "Connected to Odoo Server.";
                        PrintNLog(info);
                        return true;
                    }
                    info = "Status: " + reply.Status + ", Time: " + reply.RoundtripTime.ToString() + ", Address: " + reply.Address + "\n";
                    PrintNLog(info);
                }
                return false;   //jika tdk ada reply atau reply selain 'Success'
            }
            catch
            {
                info = "ERROR: Not connected to Odoo Server.";
                PrintNLog(info);
                return false;
            }

        }

        static void cekAbsensi()
        {
            string sql = "";

            try
            {
                OleDbConnection connAbsensi = new OleDbConnection(connStringAbsensi);
                try
                {
                    if (connAbsensi.State == ConnectionState.Closed) connAbsensi.Open();
                }
                catch (Exception ex)
                {
                    PrintNLog(ex.Message + "\n\nPress any key to close.");
                    Console.ReadLine();
                    Environment.Exit(0);
                }

                if (maxRecords == 0)
                    sql = "SELECT ";
                else
                    sql = $"SELECT TOP {maxRecords} ";

                sql = sql + " chk.userid, format(chk.checktime,'yyyy-mm-dd HH:mm:ss') as checktime, chk.checktype, u.name, u.badgenumber from checkinout " +
                    " chk inner join userinfo u on chk.userid = u.userid WHERE chk.senttoodoo is null or chk.senttoodoo = '' or chk.senttoodoo = 'ERROR' " +
                    " order by chk.checktime;";

                OleDbCommand cmdCheckIn = new OleDbCommand(sql, connAbsensi);
                OleDbDataReader rdr = cmdCheckIn.ExecuteReader();
                while (rdr.Read())
                {
                    string userId = rdr.GetInt32(0).ToString();
                    string checkTime = rdr.GetValue(1).ToString();
                    string checkType = rdr.GetString(2);
                    string userName = rdr.GetString(3);
                    string badgeNumber = rdr.GetString(4);

                    var dataHarian = new PostData
                    {
                        USERID = badgeNumber,
                        CHECKTIME = checkTime,
                        USERNAME = userName
                    };

                    var client = new RestClient();
                    var request = new RestRequest(baseURL, Method.Post);

                    request.AddHeader("Content-Type", "application/json");
                    request.AddHeader("Cookie", "session_id=d24ee0e4dab856ec627359b58a1d764f41c18cec");

                    Absensi _absensi = new Absensi();
                    _absensi.CheckInOut = dataHarian;
                    var json = JsonSerializer.Serialize(_absensi);
                    json = json.Replace("CheckInOut", "params");

                    var content = new StringContent(json, Encoding.UTF8, "application/json");
                    request.AddStringBody(json, DataFormat.Json);

                    if (cekKoneksiOdoo() == false)
                    {
                        logInfo = "Can't connect to Odoo API. Please check the Odoo URL.";
                        PrintNLog(logInfo);
                        Console.ReadLine();
                    }

                    RestResponse response = client.Execute(request);
                    string statusCode = response.StatusCode.ToString();

                    if (statusCode == "OK")
                    {
                        sql = $"UPDATE CHECKINOUT SET SenttoOdoo = 'SUCCESS' WHERE " +
                            $"userid = {userId} and format(checktime,'yyyy-mm-dd HH:mm:ss') = '{checkTime}' and checktype = '{checkType}'; ";
                        OleDbCommand cmdSent = new OleDbCommand(sql, connAbsensi);
                        cmdSent.CommandText = sql;
                        cmdSent.ExecuteNonQuery();
                    }
                    else
                    {
                        sql = $"UPDATE CHECKINOUT SET SenttoOdoo = 'ERROR' WHERE " +
                            $"userid = {userId} and format(checktime,'yyyy-mm-dd HH:mm:ss') = '{checkTime}' and checktype = '{checkType}'; ";
                        OleDbCommand cmdSent = new OleDbCommand(sql, connAbsensi);
                        cmdSent.CommandText = sql;
                        cmdSent.ExecuteNonQuery();
                    }

                    logInfo = $"USERID: {userId}, CHECKTIME: {checkTime}, NAME: {userName}, BADGENUMBER: {badgeNumber}, Sent Status: " + statusCode;
                    PrintNLog(logInfo);

                    if (Console.KeyAvailable)
                    {
                        connAbsensi.Close();
                        exit = true;
                        
                        logInfo = "User stopped at " + DateTime.Now;
                        PrintNLog(logInfo);
                        SaveLogFile();

                        Console.ReadLine();
                        Environment.Exit(0);
                    }

                    Thread.Sleep(400);
                }
                connAbsensi.Close();

                logInfo = "Finished at " + DateTime.Now;
                PrintNLog(logInfo);
                SaveLogFile();
                Console.ReadLine();
            }
            catch (Exception ex)
            {
                PrintNLog(ex.Message);
            }
        }

        static void PrintNLog(string logMessage)
        {
            Console.WriteLine(logMessage);
            completeLogText = completeLogText + logMessage + Environment.NewLine;
        }

        static void SaveLogFile()
        {
            if (completeLogText.Length == 0) return;

            string logFolder = "logs";

            if (!Directory.Exists(logFolder)) Directory.CreateDirectory(logFolder);

            var waktu = DateTime.Now.ToString("yyyyMMdd-HHmm");
            string logfile = String.Format($@"{logFolder}\Fingerprint-{waktu}.log");

            System.IO.StreamWriter SaveFile = new System.IO.StreamWriter(logfile);
            SaveFile.WriteLine(completeLogText);
            SaveFile.Close();
            completeLogText = "";
        }

    }
}
