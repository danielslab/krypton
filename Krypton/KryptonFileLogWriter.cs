using System.IO;
using Common;

namespace Krypton
{
    class KryptonFileLogWriter : IKryptonLogger
    {
        private StreamWriter _logwriter;
        readonly int _testCaseCount;
        public KryptonFileLogWriter(int testcasecnt) 
        {
            _testCaseCount = testcasecnt;
        }

        public void WriteLog(string message) 
        {
            DirectoryInfo loglocation = new DirectoryInfo(Property.ResultsDestinationPath);
            if (Directory.Exists(loglocation.FullName))
            {
                using (_logwriter = new StreamWriter(loglocation.FullName + "/" + "Log" + _testCaseCount + ".txt", true))
                {
                    _logwriter.WriteLine(message);
                }
            }
            else
            {
                string folderPath = Path.Combine(Property.IniPath, Property.LogFolder);
                using (_logwriter = new StreamWriter(folderPath + "/" + "Log" + _testCaseCount + ".txt", true))
                {
                    _logwriter.WriteLine(message);
                }
            }
        }
    }
}
