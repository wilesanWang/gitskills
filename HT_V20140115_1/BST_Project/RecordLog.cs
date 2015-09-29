using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;

namespace BST_Project
{
    class RecordLog
    {
        public static string LOG_HEADER_FORMAT = System.DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
        public static string LOG_PATH = CommonVary.logPath + "\\Log_"+DateTime.Today.ToString("yyyyMMdd") + ".txt";
        public static FileOperation fileOperation = new FileOperation(LOG_PATH);
        public static string[] LEVELSTRING = new string[]
        {   
            "ERROR",
            "WARNING",
            "INFO",
            "DEBUG",
            "SUCCESS"
        };
        public enum LOGLEVEL
        {
            LOGERROR = 0,
            LOGWARNING,
            LOGINFO,
            LOGDEBUG,
            LOGSUCCESS
        };
        LOGLEVEL _curLOGLevel;
        /// <summary>
        /// 获得日志级别
        /// </summary>
        /// <param name="level"></param>
        /// <returns></returns>
       private static string getLogLevelDesp(LOGLEVEL level)
        {
            if ((uint)level > (uint)(LOGLEVEL.LOGSUCCESS))
            {
                return "Unknow Level";
            }
            return LEVELSTRING[(uint)level];
        }
        /// <summary>
        /// 写日志
        /// </summary>
        /// <param name="context"></param>
        /// <returns></returns>
        public static int _WriteLog(LOGLEVEL level, string context)
        {
            LOG_PATH = CommonVary.logPath + "\\Log_" + DateTime.Today.ToString("yyyyMMdd") + ".txt";

            string LogLevel = getLogLevelDesp((LOGLEVEL)level);
            context = LOG_HEADER_FORMAT + " " + LogLevel + " :"+ context;
            int rc = CommonVary.RUNNING_OK;
            try
            {
                fileOperation.CreateIfNotExist(LOG_PATH);
                fileOperation.WriteLine(context);

            }
            catch (System.Exception ex)
            {
                MessageBox.Show("日志写入出错：" + context);
                return CommonVary.RUNNNING_WRONG;
            }
            return rc;
        }

        public static int _CreateNewLog()
        {
            bool a = File.Exists(LOG_PATH);
            if (!File.Exists(LOG_PATH))
            {
                _WriteLog(RecordLog.LOGLEVEL.LOGSUCCESS, "日志记录：" + DateTime.Today.ToString("yyyy－MM－dd"));
             }
            return 0;
        }
    }
}
