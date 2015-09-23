/***************************************************************************
** Application: Krypton: ThinkSys Automation Generic System
** File Name: BrowserManager.cs
** Copyright 2012, ThinkSys Software
** All rights reserved
** Developed by: ThinkSys Software 
** Description: Browser Specific Functionality file
*****************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Win32;
using System.Diagnostics;

namespace Common
{
    /// <summary>
    /// This file will take care of browser specific functionalities.
    /// </summary>

    public interface IBrowser
    {
        string validateBrowserInstalled();
        Process[] getBrowserRunningProcess();
        void clearCache();
    }

    /// <summary>
    /// Handles IE related functionalities.
    /// </summary>
    class IE : IBrowser
    {
        public string validateBrowserInstalled()
        {
            bool x = false;
            string SoftwareKey = Property.Browser_KeyString;
            try
            {
                using (RegistryKey rk = Registry.LocalMachine.OpenSubKey(SoftwareKey))
                {
                    foreach (string skName in rk.GetSubKeyNames())
                    {
                        if (skName.ToLower().Trim().Contains("ie"))
                        {
                            x = true;
                        }
                    }
                    if (x)
                        return ExecutionStatus.Pass;
                    else
                        return ExecutionStatus.Fail;
                }
            }
            catch (Exception e)
            {
                throw new KryptonException(Utility.GetCommonMsgVariable("KRYPTONERRCODE0001"));
            }
        }

        public Process[] getBrowserRunningProcess()
        {
            return Process.GetProcessesByName(Property.IE_PROCESS);
        }
        public void clearCache()
        {
            string[] s = System.IO.Directory.GetFiles(Environment.GetFolderPath(Environment.SpecialFolder.Cookies));
            foreach (string currentFile in s)
            {
                try
                {
                    System.IO.File.Delete(currentFile);
                }
                catch (Exception)
                {

                }
            }
        }
    }
    /// <summary>
    /// Handles Firefox related functionalities. 
    /// </summary>
    class Firefox : IBrowser
    {
        public string validateBrowserInstalled()
        {
            bool isFF = false;
            string SoftwareKey = Property.Browser_KeyString;
            try
            {
                using (RegistryKey rk = Registry.LocalMachine.OpenSubKey(SoftwareKey))
                {
                    foreach (string skName in rk.GetSubKeyNames())
                    {
                        if (skName.Contains("Mozilla Firefox"))
                            isFF = true;
                    }
                    if (isFF)
                        return ExecutionStatus.Pass;
                    else
                        return ExecutionStatus.Fail;
                }
            }
            catch (Exception e)
            {
                throw new KryptonException(Utility.GetCommonMsgVariable("KRYPTONERRCODE0002"));
            }
        }

        public Process[] getBrowserRunningProcess()
        {
            return Process.GetProcessesByName(Property.FF_Process);
        }

        public void clearCache()
        {
            try
            {
                string path = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
                string cachepath = path + @"\Mozilla\Firefox\Profiles\Cache";
                if (!System.IO.Directory.Exists(cachepath))
                    return;
                string[] files = System.IO.Directory.GetFiles(cachepath);
                foreach (string file in files)
                {
                    System.IO.File.Delete(file);
                }
            }
            catch (Exception)
            {

            }
        }

    }

    /// <summary>
    /// Handles Chrome related functionalities. 
    /// </summary>
    class Chrome : IBrowser
    {
        public string validateBrowserInstalled()
        {
            bool isChrome = false;
            string SoftwareKey = Property.Chrome_Specific_KeyString;
            try
            {
                using (RegistryKey rk = Registry.LocalMachine.OpenSubKey(SoftwareKey))
                {
                    foreach (string skName in rk.GetSubKeyNames())
                    {
                        if (skName.Contains(KryptonConstants.BROWSER_CHROME))
                        {
                            isChrome = true;
                            break;
                        }
                    }
                    if (isChrome)
                        return ExecutionStatus.Pass;
                    else
                        return ExecutionStatus.Fail;
                }
            }
            catch (Exception e)
            {
                throw new KryptonException(Utility.GetCommonMsgVariable("KRYPTONERRCODE0003"));
            }
        }

        public Process[] getBrowserRunningProcess()
        {
            return Process.GetProcessesByName(Property.CHROME_PROCESS);
        }
        public void clearCache()
        {
            //Code need to be written.
        }
    }

    /// <summary>
    /// Decide the browser on which execution run.
    /// </summary>

    public class BrowserManager
    {
        public static IBrowser browser;
        public static IBrowser getBrowser(string browserrun)
        {
            if (browserrun.ToLower().Equals(KryptonConstants.BROWSER_IE))
            {
                browser = new IE();
            }
            else if (browserrun.ToLower().Equals(KryptonConstants.BROWSER_FIREFOX))
            {
                browser = new Firefox();
            }
            else if (browserrun.ToLower().Equals(KryptonConstants.BROWSER_CHROME))
            {
                browser = new Chrome();
            }
            return browser;
        }

    }
}
