using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using Common;

namespace Krypton
{
    class KryptonFileLogWriter : IKryptonLogger
    {
        int testCaseCnt = 0;
        StreamWriter logwriter = null;
        public KryptonFileLogWriter(int testcasecnt) 
        {
            testCaseCnt = testcasecnt;
            DirectoryInfo loglocation = new DirectoryInfo(Property.ResultsDestinationPath);
            loglocation.Create();
            logwriter = new StreamWriter(loglocation.FullName + "/" + "Log" + testCaseCnt + ".txt");
        }
        public void WriteLog(string message) 
        {
            logwriter.WriteLine(message);
        }
    }
}
