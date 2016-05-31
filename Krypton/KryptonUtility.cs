using Common;
using System.IO;

namespace Krypton
{
    class KryptonUtility : Variables
    {
        /// <summary>
        /// Set the all file location into the vairables in the property.cs 
        /// To identify the location of all the project files.
        /// </summary>
        /// <param></param> 
        /// <returns></returns>
        internal static void SetProjectFilesPaths()
        {
            //Get Data File path
            Property.TestDataFilepath = Path.IsPathRooted(Utility.GetParameter("TestDataLocation")) ? Utility.GetParameter("TestDataLocation") : Path.Combine(Property.IniPath, Utility.GetParameter("TestDataLocation"));

            //Get DBTestData File path
            Property.DBTestDataFilepath = Path.IsPathRooted(Utility.GetParameter("DBTestDataLocation")) ? Utility.GetParameter("DBTestDataLocation") : Path.Combine(Property.IniPath, Utility.GetParameter("DBTestDataLocation"));

            //get reusable path
            Property.ReusableLocation = Path.IsPathRooted(Utility.GetParameter("DBTestDataLocation")) ? Utility.GetParameter("ReusableLocation") : Path.Combine(Property.IniPath, Utility.GetParameter("ReusableLocation"));

            //Get OR File path
            Property.ObjectRepositoryFilepath = Path.IsPathRooted(Utility.GetParameter("ORLocation")) ? Utility.GetParameter("ORLocation") : Path.GetFullPath(Path.Combine(Property.IniPath, Utility.GetParameter("ORLocation")));

            #region RecoverFromPopuppath.
            Property.RecoverFromPopupFilepath = Path.IsPathRooted(Utility.GetParameter("RecoverFromPopupLocation")) ? Utility.GetParameter("RecoverFromPopupLocation") : Path.GetFullPath(string.Concat(Property.ApplicationPath, Utility.GetParameter("RecoverFromPopupLocation")));

            //Get RecoverFromPopup File 
            Property.RecoverFromPopupFilename = Utility.GetParameter("RecoverFromPopupFileName") + ".xlsx";
            Property.MaxTimeoutForPageLoad = int.Parse(Utility.GetParameter("MaxTimeoutForPageLoad"));
            Property.MinTimeoutForPageLoad = int.Parse(Utility.GetParameter("MinTimeoutForPageLoad"));
            if (Path.IsPathRooted(Utility.GetParameter("RecoverFromBrowserLocation")))
                Property.RecoverFromBrowserFilePath = Utility.GetParameter("RecoverFromBrowserLocation");
            else
                Property.RecoverFromBrowserFilePath = Path.GetFullPath(string.Concat(Property.ApplicationPath, Utility.GetParameter("RecoverFromBrowserLocation")));

            Property.RecoverFromBrowserFilename = Utility.GetParameter("RecoverFromBrowserFileName") + ".xlsx";
            #endregion
            //Get Company logo path
            Property.CompanyLogo = Path.IsPathRooted(Utility.GetParameter("CompanyLogo")) ? Utility.GetParameter("CompanyLogo") : Path.GetFullPath(Path.Combine(Property.IniPath, Utility.GetParameter("CompanyLogo")));
            if (!File.Exists(Property.CompanyLogo))
                Property.CompanyLogo = string.Empty;

            if (Path.IsPathRooted(Utility.GetParameter("parallelRecovery")))
                Property.ParallelRecoveryFilePath = Utility.GetParameter("parallelRecovery");
            else
                Property.ParallelRecoveryFilePath = Path.GetFullPath(Path.Combine(Property.ApplicationPath, Utility.GetParameter("parallelRecovery")));
            //Database Connection String initialization
            Property.DbConnectionString = Utility.GetParameter("DBConnectionString");
            //Driver Error capture as image or html
            Property.ErrorCaptureAs = Utility.GetParameter("ErrorCaptureAs");
            //extract parallel recoverysheet name from parameter.ini
            if (Utility.GetParameter("ParallelRecoverySheetName").Trim().Length > 0)
            {
                Property.ParallelRecoverySheetName = Utility.GetParameter("ParallelRecoverySheetName") + ".xlsx";
            }
            else
            {
                Property.ParallelRecoverySheetName = "popuprecovery" + Property.ExcelSheetExtension;
            }

        }
    }
}
