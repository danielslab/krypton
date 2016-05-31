/***************************************************************************
** Application: Krypton: ThinkSys Automation Generic System
** File Name: BrowserManager.cs
** Copyright 2012, ThinkSys Software
** All rights reserved
** Developed by: ThinkSys Software 
** Description: Browser Specific Functionality file
*****************************************************************************/
using System;
using Microsoft.Win32;
using System.Diagnostics;
using System.Linq;

namespace Common
{
    /// <summary>
    /// This file will take care of browser specific functionalities.
    /// </summary>

    public interface IBrowser
    {
        string ValidateBrowserInstalled();
        Process[] GetBrowserRunningProcess();
        void ClearCache();
    }

    /// <summary>
    /// Handles IE related functionalities.
    /// </summary>
    class Ie : IBrowser
    {
        public string ValidateBrowserInstalled()
        {
            bool x = false;
            string softwareKey = Property.BrowserKeyString;
            try
            {
                using (RegistryKey rk = Registry.LocalMachine.OpenSubKey(softwareKey))
                {
                    if (rk != null)
                        foreach (string skName in rk.GetSubKeyNames())
                        {
                            if (skName.ToLower().Trim().Contains("ie"))
                            {
                                x = true;
                            }
                        }
                    if (x)
                        return ExecutionStatus.Pass;
                    return ExecutionStatus.Fail;
                }
            }
            catch (Exception e)
            {
                throw new KryptonException(Utility.GetCommonMsgVariable("KRYPTONERRCODE0001"));
            }
        }

        public Process[] GetBrowserRunningProcess()
        {
            return Process.GetProcessesByName(Property.IeProcess);
        }

        public void ClearCache()
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
                    // ignored
                }
            }
        }
    }
    /// <summary>
    /// Handles Firefox related functionalities. 
    /// </summary>
    class Firefox : IBrowser
    {
        public string ValidateBrowserInstalled()
        {
            bool isFf = false;
            string softwareKey = Property.BrowserKeyString;
            try
            {
                using (RegistryKey rk = Registry.LocalMachine.OpenSubKey(softwareKey))
                {
                    foreach (string skName in rk.GetSubKeyNames())
                    {
                        if (skName.Contains("Mozilla Firefox"))
                            isFf = true;
                    }
                    if (isFf)
                        return ExecutionStatus.Pass;
                    return ExecutionStatus.Fail;
                }
            }
            catch (Exception e)
            {
                throw new KryptonException(Utility.GetCommonMsgVariable("KRYPTONERRCODE0002"));
            }
        }

        public Process[] GetBrowserRunningProcess()
        {
            return Process.GetProcessesByName(Property.FfProcess);
        }

        public void ClearCache()
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
                // ignored
            }
        }

    }

    /// <summary>
    /// Handles Chrome related functionalities. 
    /// </summary>
    class Chrome : IBrowser
    {
        public string ValidateBrowserInstalled()
        {
            bool isChrome = false;
            string softwareKey = Property.ChromeSpecificKeyString;
            try
            {
                using (RegistryKey rk = Registry.LocalMachine.OpenSubKey(softwareKey))
                {
                    if (rk != null && rk.GetSubKeyNames().Any(skName => skName.Contains(KryptonConstants.BROWSER_CHROME)))
                    {
                        isChrome = true;
                    }
                    if (isChrome)
                        return ExecutionStatus.Pass;
                    return ExecutionStatus.Fail;
                }
            }
            catch (Exception)
            {
                throw new KryptonException(Utility.GetCommonMsgVariable("KRYPTONERRCODE0003"));
            }
        }

        public Process[] GetBrowserRunningProcess()
        {
            return Process.GetProcessesByName(Property.ChromeProcess);
        }

        public void ClearCache()
        {
            //Code need to be written.
        }
    }

    /// <summary>
    /// Decide the browser on which execution run.
    /// </summary>

    public class BrowserManager
    {
        public static IBrowser Browser;
        public static IBrowser GetBrowser(string browserrun)
        {
            if (browserrun.ToLower().Equals(KryptonConstants.BROWSER_IE))
            {
                Browser = new Ie();
            }
            else if (browserrun.ToLower().Equals(KryptonConstants.BROWSER_FIREFOX))
            {
                Browser = new Firefox();
            }
            else if (browserrun.ToLower().Equals(KryptonConstants.BROWSER_CHROME))
            {
                Browser = new Chrome();
            }
            return Browser;
        }

    }
}
