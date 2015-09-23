using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration.Install;
using System.Linq;
using System.IO;
using System.Windows.Forms;


namespace Krypton
{
    [RunInstaller(true)]
    public partial class Installer1 : System.Configuration.Install.Installer
    {
        public Installer1()
        {
            InitializeComponent();
        }

        private static string _msgAccessDatabaseInstall1 = "Please Install Microsoft Access Database Engine 32-bit.";
        private static string _msgAccessDatabaseInstall2 = "Uninstall 64-bit Version of Access Database Engine.Then Download and Install the 32-bit version of Access Database.";

        public override void Install(IDictionary stateSaver)
        {
            if (System.Environment.Is64BitOperatingSystem && !Directory.Exists(@"C:\Program Files (x86)\Common Files\microsoft shared\OFFICE14") && !Directory.Exists(@"C:\Program Files\Common Files\Microsoft Shared\OFFICE14"))
            {
                throw new System.Configuration.Install.InstallException(_msgAccessDatabaseInstall1);
                //throw new Exception("Please Install Microsoft Access Database Engine 32-bit.");
            }
            else if (System.Environment.Is64BitOperatingSystem && Directory.Exists(@"C:\Program Files\Common Files\microsoft shared\OFFICE14"))
            {
                throw new System.Configuration.Install.InstallException(_msgAccessDatabaseInstall2);
                //MessageBox.Show(_msgAccessDatabaseInstall2);
                //throw new Exception("Uninstall 64-bit Version of Access Database Engine.Then Download and Install the 32-bit version of Access Database.");
            }
        }

    }
}
