using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using MySql.Data.MySqlClient;

namespace BST_Project
{
    class CommonVary
    {
        public static int dataLength = Marshal.SizeOf(MainForm.sendMapData);//定义发送数据长度
        public static int RUNNING_OK = 0;
        public static int RUNNNING_WRONG = -1;
        public static string logPath = System.Environment.CurrentDirectory.ToString();
        public static string configPath = logPath + @"\config.ini";
        public static System.Int16[] data = new Int16[dataLength/2];


        public static string PLCAddress = "";
        public static string ServerAddress = "";
        public static int PLCPort = 0;
        public static int ServerPort = 0;

        public static int PCFromPLCPort = 0;
        public static int PCFromServerPort = 0;

        public static string myCOMPort = "";
        public static string readCOMPort = "";
        public static int COMBaunRate = 0;

        public static int SaveDataInterval = 0;
        public static string configFile;


        public static string cbx4Item1 = "料条";//"Stripe";
        public static string cbx4Item2 = "边";//"Edge";
        public static string cbx5Item1 = "单相机";//"Single CCD";
        public static string cbx5Item2 = "双相机";//"Dual CCD";

        public static string cbx3Item1 = "单物料";//"Single Web";
        public static string cbx3Item2 = "双物料";//"Dual Web";
        public static string cbx6Item1 = "单工位";//"Single Pos";
        public static string cbx6Item2 = "双工位";//"Dual Pos";

        public static string mysqlConnectStr = @"Database=bst_project;Data Source=localhost;User Id=root;Password=root;Allow Zero Datetime=true";
        public static MySqlConnection conn = new MySqlConnection(mysqlConnectStr); 
        public static int OpenDataConnection()
        {
            try
            {
                //数据库连接  
                if (conn.State != System.Data.ConnectionState.Open)
                {
                    conn.Open();
                }
            }
            catch (MySql.Data.MySqlClient.MySqlException ex)
            {
                MessageBox.Show("连接数据库出错");
                RecordLog._WriteLog(RecordLog.LOGLEVEL.LOGERROR, "连接数据库出错！\n信息：" + ex.Message);
                return 1;
            }

            return 0;
            
        }

        public static int CloseDataConnection()
        {
            try
            {
                if (conn.State != System.Data.ConnectionState.Closed)
                {
                    conn.Close();
                }
            }
            catch (MySql.Data.MySqlClient.MySqlException ex)
            {
                MessageBox.Show("关闭数据库出错");
                RecordLog._WriteLog(RecordLog.LOGLEVEL.LOGERROR, "关闭数据库出错！\n信息：" + ex.Message);
                return 1;
            }

            return 0;

        }

        public static int ReadConfig()
        {
            if(!System.IO.File.Exists(configPath))
            {
                RecordLog._WriteLog(RecordLog.LOGLEVEL.LOGERROR,"配置文件不存在");
                MessageBox.Show("配置文件不存在！");
                return CommonVary.RUNNNING_WRONG;
            }
            else
            {
                FileStream fsRead = new FileStream(configPath, FileMode.Open, FileAccess.Read);
                StreamReader sr = new StreamReader(fsRead);

                configFile = sr.ReadToEnd();
                string[] lineInfo = configFile.Split('\n');
                string[] key_value = new string[2];
                try
                {
                    for(int i = 0 ;i< lineInfo.Length;i++)
                    {
                        key_value = lineInfo[i].Split(':');
                        switch (key_value[0])
                        {
                            case "PLCAddress":
                                PLCAddress = key_value[1].Trim();break;
                            case "PLCPort":
                                PLCPort = int.Parse(key_value[1].Trim());break;
                            case "PCFromPLCPort":
                                PCFromPLCPort = int.Parse(key_value[1].Trim());break;
                            case "SaveDataInterval":
                                SaveDataInterval = int.Parse(key_value[1].Trim()) * 1000;break;
                            case "WriteCOMPort":
                                myCOMPort = key_value[1].Trim();break;
                            case "COMBaunRate":
                                COMBaunRate = int.Parse(key_value[1].Trim());break;
                           

                            default:
                                RecordLog._WriteLog(RecordLog.LOGLEVEL.LOGERROR,"配置文件中有不规则信息");
                                MessageBox.Show("配置文件中有不规则信息");
                                break;
                        }
                        
                    
                    }
                }
                catch(Exception ex)
                {
                    sr.Close();
                    fsRead.Close();
                    RecordLog._WriteLog(RecordLog.LOGLEVEL.LOGERROR,"读取配置文件出错");
                    MessageBox.Show("读取配置文件出错,修改配置文件并重启");
                    return CommonVary.RUNNNING_WRONG;
                }
                if (PLCAddress == "" || PLCPort == 0 || PCFromPLCPort == 0 || SaveDataInterval == 0)
                {
                    RecordLog._WriteLog(RecordLog.LOGLEVEL.LOGERROR, "读取配置文件出错");
                    MessageBox.Show("读取配置文件出错,修改配置文件并重启");
                    sr.Close();
                    fsRead.Close();
                    return CommonVary.RUNNNING_WRONG;
                }
                else
                {
                    sr.Close();
                    fsRead.Close();
                    return CommonVary.RUNNING_OK;
                }
            }
        }

        public static int WriteConfig(string type,string value,string updateInfo)
        {
            FileStream fsWrite = new FileStream(configPath, FileMode.Open, FileAccess.Write);
            StreamWriter sw = new StreamWriter(fsWrite);
            
            try
            {
                configFile = configFile.Replace(type + ":" + value, type + ":" + updateInfo);
                sw.Write(configFile);
                sw.Close();
                fsWrite.Close();
                if (type == "SaveDataInterval" || type == "WriteCOMPort")
                    MessageBox.Show("配置文件修改成功");
                else
                    MessageBox.Show("配置文件修改成功，请重启测宽软件！");
                ReadConfig();
                return CommonVary.RUNNING_OK;
            }
            catch(Exception ex)
            {
                sw.Close();
                fsWrite.Close();
                RecordLog._WriteLog(RecordLog.LOGLEVEL.LOGERROR, "配置文件写入出错");
                MessageBox.Show("配置文件写入出错");
                return CommonVary.RUNNNING_WRONG;
            }


        }
    }
}
