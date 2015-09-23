/***************************************************************************
** Application: Krypton: ThinkSys Automation Generic System
** File Name: Common.KRYPTONException.cs
** Copyright 2012, ThinkSys Software
** All rights reserved
** Developed by: ThinkSys Software 
** Description: Custom Exception handler
*****************************************************************************/
using System;
using System.Runtime.Serialization;
using System.IO;

namespace Common
{
    [Serializable]
    public sealed class KryptonException : Exception
    {
        private string stringInfo;

        public KryptonException() : base() { }
        public KryptonException(string message) : base(message) { }

        private KryptonException(SerializationInfo info, StreamingContext context)
            : base(info, context)
        {
            stringInfo = info.GetString("StringInfo");
        }

        public KryptonException(string message, string stringInfo)
            : this(message)
        {
            this.stringInfo = stringInfo;
        }

        public string StringInfo
        {
            get { return stringInfo; }
        }


        public override void GetObjectData(SerializationInfo info, StreamingContext context)
        {
            info.AddValue("StringInfo", stringInfo);

            base.GetObjectData(info, context);
        }

        public override string Message
        {
            get
            {
                string message = base.Message;
                if (stringInfo != null)
                {
                    message += ":\t" + stringInfo;
                }
                return message;
            }
        }

        // for Logging exceptions thrown in a Text File 
        // Add "Common.KryptonException.writeexception(e);" in each catch case
        public static void writeexception(Exception e)
        {
                string s =Common.Utility.GetParameter("TestCaseId");
                string folderPath = Path.Combine(Common.Property.IniPath, Common.Property.logFolder);
                if (!Directory.Exists(folderPath))
                {
                    Directory.CreateDirectory(folderPath);
                }
                string filepath = Path.Combine(folderPath, Common.Property.CURRENTTIME);
                StreamWriter obj = File.AppendText(filepath);
                obj.WriteLine("TestCaseId: " + s + "Exception : " + e.StackTrace.ToString() + "Message: " + e.Message);
                obj.Close();
        }

    }
}
