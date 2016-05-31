/***************************************************************************
** Application: Krypton: ThinkSys Automation Generic System
** File Name: Common.Validate.cs
** Copyright 2012, ThinkSys Software
** All rights reserved
** Developed by: ThinkSys Software 
** Description: Handle validation Process.
*****************************************************************************/

using System;
using Microsoft.Win32;

namespace Common
{

    public interface validation
    {
        string ValidationProcess();
    }

    /// <summary>
    /// Validate Selenium related required software installations.
    /// </summary>
    class ValidateSeleniumSetup : validation
    {
        public string ValidationProcess()
        {
            Property.StepComments = "Selenium driver installed?";
            string JSoftware = CommonValidationMethods.ValidateJavaSoftwareInstalled();
            string Jversion = CommonValidationMethods.validateJavaVersion();
            string BrowserInstalled;
            BrowserInstalled = BrowserManager.Browser.ValidateBrowserInstalled();

            if (JSoftware.Equals(ExecutionStatus.Fail) || Jversion.Equals(ExecutionStatus.Fail) || BrowserInstalled.Equals(ExecutionStatus.Fail))
                return ExecutionStatus.Fail;
            else
                return ExecutionStatus.Pass;

        }
    }
    /// <summary>
    /// Validate QTP related required software installations.
    /// </summary>
    class ValidateQTPSetup : validation
    {
        public string ValidationProcess()
        {
            Property.StepComments = "QTP driver installed?";
            // QTP related software validation will add here.
            return ExecutionStatus.Pass;
        }
    }

    /// <summary>
    /// Decide DRIVER at run time.
    /// </summary>
    public class Validate
    {
        public static validation validate;
        public static validation GetDriverValidation(string driver)
        {
            if (driver.Equals(Property.Selenium))
            {
                validate = new ValidateSeleniumSetup();
            }
            else if (driver.Equals(Property.Qtp))
            {
                validate = new ValidateQTPSetup();
            }
            return validate;
        }
    }

    /// <summary>
    /// Class contains actual functionality for validation process.
    /// </summary>
    class CommonValidationMethods
    {

        /// <summary>
        /// Validate whether Java is installed in execution system.
        /// </summary>
        /// <returns></returns>
        public static string ValidateJavaSoftwareInstalled()
        {
            int flagJava = 0;
            string softwareKey = Property.JavaKeyString;
            try
            {
                using (RegistryKey rk = Registry.LocalMachine.OpenSubKey(softwareKey))
                {
                    foreach (string skName in rk.GetSubKeyNames())
                    {
                        if (skName.Equals(Property.JavaRunTimeEnvironmentString))
                            flagJava = 1;
                    }
                    if (flagJava == 1)
                        return ExecutionStatus.Pass;
                    return ExecutionStatus.Fail;
                }
            }
            catch (Exception e)
            {
                throw new KryptonException(Utility.GetCommonMsgVariable("KRYPTONERRCODE0004").Replace("{MSG}", e.Message));
            }
        }

        /// <summary>
        /// Validate whether the version of java is heigher or equals to 1.5.
        /// </summary>
        /// <returns>Return String with status Pass OR Fail.</returns>
        public static string validateJavaVersion()
        {
            bool isAbove = false;
            string SoftwareKey = Property.JavaRunTimeEnvironmentKeyString;
            try
            {
                using (RegistryKey rk = Registry.LocalMachine.OpenSubKey(SoftwareKey))
                {

                    foreach (string skName in rk.GetSubKeyNames())
                    {
                        string sk = skName.Substring(2, 1);
                        if (int.Parse(sk) >= 5)
                        {
                            isAbove = true;
                        }
                    }
                    if (isAbove)
                        return ExecutionStatus.Pass;
                    else
                        return ExecutionStatus.Fail;
                }

            }
            catch (Exception e)
            {
                throw new KryptonException(Utility.GetCommonMsgVariable("KRYPTONERRCODE0005").Replace("{MSG}", e.Message));
            }
        }

    }
}
