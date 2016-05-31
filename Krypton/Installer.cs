using System;
using System.Collections;
using System.ComponentModel;
using System.Diagnostics;
using System.Windows.Forms;


namespace Krypton
{
    [RunInstaller(true)]
    public partial class Installer : System.Configuration.Install.Installer
    {
        public Installer()
        {
            InitializeComponent();
        }

        protected override void OnAfterInstall(IDictionary savedState)
        {
            if (Environment.Is64BitOperatingSystem)
            {
                try
                {
                    Process reg = new Process
                    {
                        StartInfo =
                        {
                            FileName = "regsvr32.exe",
                            Arguments = @"C:\Krypton\Output\AutoItX3_x64.dll",
                            UseShellExecute = false,
                            CreateNoWindow = true,
                            RedirectStandardOutput = true
                        }
                    };
                    //This file registers .dll files as command components in the registry.
                    reg.Start();
                    reg.WaitForExit();
                    reg.Close();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            base.OnAfterInstall(savedState);
        }
    }
}
