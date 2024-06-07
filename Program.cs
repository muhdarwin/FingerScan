using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Net.Http;
using System.Data;
using System.Data.OleDb;
using System.Security.Cryptography;
using System.Text.Json;
using System.Reflection;
using RestSharp;
using System.Diagnostics;
using static System.Net.Mime.MediaTypeNames;
using System.Web.UI.WebControls;
using System.Net.NetworkInformation;
using System.Web.Security;
using System.Data.Common;

namespace ADR_FPWare
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

    public class ResponseData
    {
        public string jsonrpc { get; set; }
        public string id { get; set; }
        public class results
        {
            public string error { get; set; }
            public string code { get; set; }
        }
    }

    internal class Program
    {
        static string connStringAbsensi = "";
        static string dbLocation = "";
        static string connStringHistory = "";
        static string baseURL = "";
        static Int32 delayTime = 0;
        static string logInfo = "";
        static string completeLogText = "";
        static int readLastNDays = 0;
        static volatile bool exit = false;
        static string tempMDB = "db";
        static string currentTime = DateTime.Now.ToString("yyyyMMddHHmmss");
        static string destinationFile = $@"{tempMDB}\temp_{currentTime}.mdb";
        static string destinationLDB = $@"{tempMDB}\temp_{currentTime}.ldb";
        static string[] schedules;

        public static void Main(string[] args)
        {
            logInfo = "Application started: " + DateTime.Now;
            PrintNLog(logInfo);

            if (!Directory.Exists(tempMDB)) Directory.CreateDirectory(tempMDB);
            readConfigFile();

            while (!exit)
            {
                Skedul();

                if (Console.KeyAvailable && Console.ReadKey(true).Key == ConsoleKey.Q)
                {
                    exit = true;

                    logInfo = "User stopped at " + DateTime.Now + "\n\nPress any key to close.";
                    PrintNLog(logInfo);
                    SaveLogFile();

                    File.Delete(destinationFile);
                    Console.ReadLine();
                    Environment.Exit(0);
                }
            }

        }

        private static void Skedul()
        {
            OleDbConnection connHistory = new OleDbConnection(connStringHistory);
            connHistory.Open();
            string checkDateTime = "";
            string sql = "SELECT * from LASTPROCESSED";
            OleDbCommand cmdHistory = new OleDbCommand(sql, connHistory);
            OleDbDataReader rdrHistory = cmdHistory.ExecuteReader();
            if (rdrHistory.Read())     // jika sdh ada datanya
            {
                checkDateTime = rdrHistory.GetValue(0).ToString();
            }
            rdrHistory.Close();

            string nextSchedule = "";
            foreach (string item in schedules)
            {
                string waktu = DateTime.Now.ToString("yyyy-MM-dd") + $" {item}";
                if (DateTime.Now <= Convert.ToDateTime(waktu))
                {
                    nextSchedule = waktu;
                    break;
                }
            }
            if (nextSchedule == "")
                nextSchedule = DateTime.Now.AddDays(1).ToString("yyyy-MM-dd") + $" {schedules[0]}";

            while (true)
            {
                if (DateTime.Now >= Convert.ToDateTime(nextSchedule))
                {
                    Proses();
                    break;
                }
            }

            sql = "SELECT * from LASTPROCESSED";
            cmdHistory = new OleDbCommand(sql, connHistory);
            rdrHistory = cmdHistory.ExecuteReader();
            if (rdrHistory.Read())
            {
                sql = $"UPDATE LASTPROCESSED SET [LAST] = '{nextSchedule}'";
            }
            else
            {
                sql = $"INSERT INTO LASTPROCESSED VALUES ('{nextSchedule}')";
            }
            rdrHistory.Close();
            cmdHistory = new OleDbCommand(sql, connHistory);
            cmdHistory.ExecuteNonQuery();
            connHistory.Close();
        }

        private static void Proses()
        {
            logInfo = $"Process started: {DateTime.Now}\n";
            PrintNLog(logInfo);

            if (cekKoneksiAbsensi())
                cekAbsensi();

            logInfo = $"Process finished: {DateTime.Now}\n";
            PrintNLog(logInfo);
            SaveLogFile();
        }

        static void readConfigFile()
        {
            StreamReader sr = new StreamReader("ADR-FPWare.ini");
            string line = sr.ReadLine();

            while (line != null)
            {
                string kar = line.Substring(0, 1);
                if ("#'".Contains(kar) == false)
                {
                    var split = line.Split('=');
                    string keyName = split[0].Replace(" ", "").ToLower();
                    string value = split[1].Trim();

                    switch (keyName)
                    {
                        case "scheduledtimes":
                            schedules = value.ToString().Replace(" ", "").Split(',');
                            break;
                        case "databaselocation":
                            dbLocation = value;
                            break;
                        case "historydatabaselocation":
                            connStringHistory = @"provider=Microsoft.Jet.OLEDB.4.0; Data Source=" + value;
                            break;
                        case "odoourl":
                            baseURL = value.ToString().Replace(@"\", "/");
                            break;
                        case "delayeachrecordevery(inmiliseconds)":
                            delayTime = Convert.ToInt32(value);
                            break;
                        case "readlastndays":
                            readLastNDays = Convert.ToInt16(value);
                            break;
                        default: break;
                    }
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

                for (int i = 0; i <= 60; i++)
                {
                    PingReply reply = myPing.Send(url, 1000);
                    if (reply != null && reply.Status.ToString() == "Success")
                    {
                        return true;
                    }
                    info = "Status: " + reply.Status + ", Time: " + reply.RoundtripTime.ToString() + ", Address: " + url + "\n";
                    PrintNLog(info);
                }
                return false;   //jika tdk ada reply atau reply selain 'Success'
            }
            catch
            {
                info = "ERROR: Can't connect to Odoo API.";
                PrintNLog(info);
                return false;
            }
        }

        static bool cekKoneksiAbsensi()
        {
            string sourceFile = dbLocation;

            try
            {
                if (!File.Exists(sourceFile))
                {
                    PrintNLog($"File {sourceFile} not exists.");
                    return false;
                }

                if (File.Exists(destinationLDB)) File.Delete(destinationLDB);
                if (File.Exists(destinationFile)) File.Delete(destinationFile);

                File.Copy(sourceFile, destinationFile);

                connStringAbsensi = $"provider=Microsoft.Jet.OLEDB.4.0; Data Source={destinationFile}";
                OleDbConnection connAbsensi = new OleDbConnection(connStringAbsensi);

                connAbsensi.Open();
                connAbsensi.Close();
                return true;
            }
            catch (IOException iox)
            {
                PrintNLog(iox.Message);
                return false;
            }
            catch (Exception ex)
            {
                PrintNLog(ex.Message + ".");
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
                    PrintNLog(ex.Message + ".\n");
                }

                OleDbConnection connHistory = new OleDbConnection(connStringHistory);
                try
                {
                    if (connHistory.State == ConnectionState.Closed) connHistory.Open();
                }
                catch (Exception ex)
                {
                    PrintNLog(ex.Message + "\n.");
                }

                sql = "SELECT chk.userid, format(chk.checktime,'yyyy-mm-dd HH:mm:ss') as checktime, chk.checktype, u.name, u.badgenumber from" +
                    " checkinout chk inner join userinfo u on chk.userid = u.userid";

                bool forTheFirstTime = false;   // set to true if this is the first time running

                if (forTheFirstTime)
                    sql = sql + $" WHERE format(chk.checktime,'yyyy-mm-dd') >= '2024-01-01' ";
                else
                {
                    if (readLastNDays > 0)  // for production: readLastNDays = 5
                    {
                        sql = sql + $" WHERE checktime >= dateadd('d', -{readLastNDays}, now())";
                    }
                }
                sql = sql + " ORDER BY chk.checktime;";

                OleDbCommand cmdCheckIn = new OleDbCommand(sql, connAbsensi);
                OleDbDataReader rdr = cmdCheckIn.ExecuteReader();
                while (rdr.Read())
                {
                    string userId = rdr.GetInt32(0).ToString();
                    string checkTime = rdr.GetValue(1).ToString();
                    string checkType = rdr.GetString(2);
                    string userName = rdr.GetString(3);
                    string badgeNumber = rdr.GetString(4);

                    sql = $"SELECT userid from DATASENT WHERE userid = {userId} AND format(checktime,'yyyy-mm-dd HH:mm:ss') = '{checkTime}'" +
                        $" AND checkType = '{checkType}' AND sentstatus = 'SUCCESS' ";
                    OleDbCommand cmdHistory = new OleDbCommand(sql, connHistory);
                    OleDbDataReader rdrHistory = cmdHistory.ExecuteReader();
                    if (rdrHistory.Read())     // jika sdh ada datanya
                    {
                        rdrHistory.Close();
                        continue;
                    }
                    rdrHistory.Close();

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
                        //logInfo = "Press any key to try again...\n";
                        //PrintNLog(logInfo);
                        //Console.ReadLine();
                    }

                    RestResponse response = client.Execute(request);
                    //var json2 = JsonSerializer.Deserialize<ResponseData>(response.Content);  // tdk digunakan

                    bool isTesting = false;
                    int errorPos = -1; // success by default 

                    if (!isTesting)
                    {
                        string responseText = response.Content;
                        errorPos = responseText.IndexOf("error");
                    }

                    string sentStatus = "";
                    if (errorPos < 0)
                    {
                        sentStatus = "SUCCESS";
                        sql = $"SELECT userid from DATASENT WHERE userid = {userId} AND format(checktime,'yyyy-mm-dd HH:mm:ss') = '{checkTime}'" +
                            $" AND checkType = '{checkType}' ";
                        cmdHistory = new OleDbCommand(sql, connHistory);
                        rdrHistory = cmdHistory.ExecuteReader();
                        if (rdrHistory.Read())     // jika sdh ada datanya
                        {
                            sql = $"UPDATE DATASENT SET SENTSTATUS = '{sentStatus}', SENTDATE = NOW() WHERE " +
                                $"userid = {userId} and format(checktime,'yyyy-mm-dd HH:mm:ss') = '{checkTime}' and checktype = '{checkType}'; ";
                        }
                        else
                        {
                            sql = $"INSERT INTO DATASENT (USERID,CHECKTIME,CHECKTYPE,SENTSTATUS) VALUES ('{userId}','{checkTime}','{checkType}','{sentStatus}')";
                        }
                        rdrHistory.Close();

                        OleDbCommand cmdSent = new OleDbCommand(sql, connHistory);
                        cmdSent.CommandText = sql;
                        cmdSent.ExecuteNonQuery();
                    }
                    else
                    {
                        sentStatus = "ERROR employee not found.";
                    }

                    logInfo = $"USERID: {userId}, CHECKTIME: {checkTime}, NAME: {userName}, BADGENUMBER: {badgeNumber}, Sent Status: {sentStatus}";
                    PrintNLog(logInfo);

                    if (Console.KeyAvailable && Console.ReadKey(true).Key == ConsoleKey.Q)
                    {
                        connAbsensi.Close();
                        connHistory.Close();
                        exit = true;

                        logInfo = "User stopped at " + DateTime.Now + "\n\nPress any key to close.";
                        PrintNLog(logInfo);
                        SaveLogFile();

                        connAbsensi.Close();
                        connHistory.Close();
                        File.Delete(destinationFile);
                        Console.ReadLine();
                        Environment.Exit(0);
                    }

                    Thread.Sleep(delayTime);
                }
                connAbsensi.Close();
                connHistory.Close();
                File.Delete(destinationFile);
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
            string logfile = String.Format($@"{logFolder}\ADR-{waktu}.log");

            System.IO.StreamWriter SaveFile = new System.IO.StreamWriter(logfile);
            SaveFile.WriteLine(completeLogText);
            SaveFile.Close();
            completeLogText = "";
        }

    }
}
