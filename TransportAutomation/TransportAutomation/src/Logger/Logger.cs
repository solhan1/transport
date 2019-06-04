using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TransportAutomation.src.Logger
{
    class Logger
    {
        public string logFilePath { get; set; }
        public Logger(string logFilePath)
        {
            this.logFilePath = logFilePath;
        }

        public void Append(string text)
        {
            if (!File.Exists(logFilePath))
            {
                /*File.Create(logFilePath);
                TextWriter tw = new StreamWriter(logFilePath);
                tw.WriteLine(text);
                tw.Close();*/
            }
            else if (File.Exists(logFilePath))
            {
                //File.AppendAllLines(logFilePath, new[] { text });
            }
        }
    }
    
}
