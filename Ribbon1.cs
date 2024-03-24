using Gone_Phishing.Properties;
using Microsoft.Office.Core;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;


namespace Gone_Phishing
{
    [ComVisible(true)]
    public class Ribbon1 : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;
        public Ribbon1()
        {
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("Gone_Phishing.Ribbon1.xml");
        }

        #endregion

        #region Ribbon Callbacks
        //Create callback methods here. For more information about adding callback methods, visit https://go.microsoft.com/fwlink/?LinkID=271226

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        public Bitmap ButtonImage(IRibbonControl control)
        {
            return Resources._80_removebg_preview_1_;
        }

        public void OnButtonClick(object sender)
        {
            ForwardSelectedEmail();
        }

        public string ReadFromRegistry(string keyPath, string valueName)
        {
            try
            {
                // Open the registry key with read access and RegistryView.Default
                using (RegistryKey key = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Default).OpenSubKey(keyPath))
                {
                    if (key != null)
                    {
                        // Read the value from the registry
                        object value = key.GetValue(valueName);


                        // Check if the value is not null
                        if (value != null)
                        {
                            return value.ToString();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show($"Error reading from registry: {ex.Message}");
            }

            return null;
        }

        public void ForwardSelectedEmail()
        {
            string registryKeyPath = null;
            if (File.Exists(@"C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE"))
            {
               registryKeyPath = @"Software\Microsoft\Office\Outlook\Addins\Unisys.GonePhishing";
            }
            else if (File.Exists(@"C:\Program Files (x86)\Microsoft Office\Office16\OUTLOOK.EXE"))
            {
                registryKeyPath = @"Software\WOW6432Node\Microsoft\Office\Outlook\Addins\Unisys.GonePhishing";
            }
            string emailAddress = ReadFromRegistry(registryKeyPath, "ReportTo");
            string prefix = ReadFromRegistry(registryKeyPath, "Prefix");

            Outlook.Explorer explorer = Globals.ThisAddIn.Application.ActiveExplorer();

            if (explorer.Selection.Count == 0)
            {
                System.Windows.Forms.MessageBox.Show("Please select an email to forward.", "No Email Selected", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Warning);
            }
            else if (explorer.Selection.Count == 1 && explorer.Selection[1] is Outlook.MailItem)
            {
                Outlook.MailItem selectedMail = explorer.Selection[1] as Outlook.MailItem;
                DialogResult result = MessageBox.Show($"Do you want to forward\n{selectedMail.Subject}\nto {emailAddress}\nand move to junk?", "Confirmation", MessageBoxButtons.OKCancel);
                
                if (result == DialogResult.OK)
                {
                    try
                    {
                        Outlook.MailItem forwardMail = selectedMail.Forward();
                        forwardMail.Recipients.Add(emailAddress);
                        forwardMail.Subject = prefix + selectedMail.Subject;
                        forwardMail.Send();

                        Outlook.MAPIFolder junkFolder = explorer.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderJunk);
                        selectedMail.Move(junkFolder);

                    }
                    catch (Exception ex)
                    {
                        System.Windows.Forms.MessageBox.Show($"{ex.Message}", "Error", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);

                    }
                }
                
            }
            else if (explorer.Selection.Count > 1)
            {
                System.Windows.Forms.MessageBox.Show("Please only forward one email", "Too Many Emails Selected", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Warning);
            }

        }

        #endregion

        #region Helpers

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion
    }
}
