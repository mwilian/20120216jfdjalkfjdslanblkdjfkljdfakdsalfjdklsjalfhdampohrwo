using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration.Install;
using System.Linq;
using Microsoft.Win32;


namespace TVCAddinCustomAction
{
    [RunInstaller(true)]
    public partial class MyInstallActions : Installer
    {
        string _appRegKeyPath = "Software\\Microsoft\\Office\\11.0\\User Settings\\QDAddin";
        string _assemblyCodeGroupName = "";
        public MyInstallActions()
        {
            InitializeComponent();
        }
        public override void Install(IDictionary stateSaver)
        {
            base.Install(stateSaver);
            //GetCodeGroupName(stateSaver);
            IncrementCount();
        }

        private void GetCodeGroupName(IDictionary stateSaver)
        {
            _assemblyCodeGroupName = this.Context.Parameters["assemblyCodeGroupName"];

            if (String.IsNullOrEmpty(_assemblyCodeGroupName))
                throw new InstallException("Cannot set the registry. The specified assembly code group name is not valid.");
            if (stateSaver == null)
                throw new ArgumentNullException("stateSaver");
            _appRegKeyPath = string.Format(_appRegKeyPath, _assemblyCodeGroupName);
        }
        public override void Uninstall(IDictionary savedState)
        {
            base.Uninstall(savedState);
            //GetCodeGroupName(stateSaver);
            IncrementCount();
            RegisterDeleteInstruction();
        }
        private void IncrementCount()
        {
            RegistryKey appKey = Registry.LocalMachine.OpenSubKey(_appRegKeyPath, true);
            object count = appKey.GetValue("Count");
            if (count == null)
                appKey.SetValue("Count", 1);
            else
                appKey.SetValue("Count", Convert.ToInt32(count) + 1);
            appKey.Close();
        }
        private void RemoveDeleteInstruction()
        {
            RegistryKey appKey = Registry.LocalMachine.OpenSubKey(_appRegKeyPath, true);
            RegistryKey deleteKey = appKey.OpenSubKey("DELETE", false);
            if (deleteKey != null)
            {
                deleteKey.Close();
                appKey.DeleteSubKey("DELETE");
            }
            appKey.Close();
        }
        private void RegisterDeleteInstruction()
        {
            RegistryKey appKey = Registry.LocalMachine.OpenSubKey(_appRegKeyPath, true);

            appKey.CreateSubKey("Delete\\Software\\Microsoft\\Office\\Excel\\AddIns\\QDAddin"); //+ _assemblyCodeGroupName

            appKey.Close();
        }
    }
}
