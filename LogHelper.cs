using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelMapWord
{
    internal class LogHelper
    {
        private static string _logPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory,"log.txt");

        public static void DeleteLogFile()
        {
            File.Delete(_logPath);
        }

        public static void Log(string message)
        {
            try
            {
                File.AppendAllText(_logPath, message);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
    }
}
