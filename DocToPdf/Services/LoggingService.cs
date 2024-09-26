using log4net;
using log4net.Appender;
using log4net.Config;
using log4net.Repository.Hierarchy;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocToPdf.Services
{
    public enum LogLevel
    {
        Info = 0,
        Debug = 1,
        Warn = 2,
        Error = 3,
        Fatal = 4
    }
    public class LoggingService
    {
        private static ILog? logging { get; set; }
        public LoggingService() { } 

        public static void LoggingInit()
        {
            XmlConfigurator.Configure(new System.IO.FileInfo("log4net.config"));
            logging = LogManager.GetLogger("Logger");
            Logger("---------- DocToPdf Start ----------", LogLevel.Info);
        }
        public static void Logger(string Log, LogLevel logLevel)
        {
            if (logging == null) return;

            switch (logLevel)
            {
                case LogLevel.Info:  logging.Info(Log);  break;
                case LogLevel.Debug: logging.Debug(Log); break;
                case LogLevel.Warn:  logging.Warn(Log);  break;
                case LogLevel.Error: logging.Error(Log); break;
                case LogLevel.Fatal: logging.Fatal(Log); break;
            }
        }        
    }
}
