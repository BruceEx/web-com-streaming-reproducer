using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace ExcelRtdTest
{
    internal class Logger
    {
        static Logger _instance = new Logger();

        public static void Log(string msg)
        {
            _instance.WriteMsg(msg);
        }

        private void WriteMsg(string msg)
        {
            // Write to the Visual Studio debugger
            Debug.WriteLine(msg);

            // Get the path of the executing assembly
            string path = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);

            // Create the full file path
            string filePath = Path.Combine(path, "log.txt");

            // Write to the local file
            File.AppendAllText(filePath, msg + Environment.NewLine);
        }
    }
}
