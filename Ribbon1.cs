using Gone_Phishing.Properties;
using Microsoft.Office.Core;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;


// TODO:  Follow these steps to enable the Ribbon (XML) item:

// 1: Copy the following code block into the ThisAddin, ThisWorkbook, or ThisDocument class.

//  protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
//  {
//      return new Ribbon1();
//  }

// 2. Create callback methods in the "Ribbon Callbacks" region of this class to handle user
//    actions, such as clicking a button. Note: if you have exported this Ribbon from the Ribbon designer,
//    move your code from the event handlers to the callback methods and modify the code to work with the
//    Ribbon extensibility (RibbonX) programming model.

// 3. Assign attributes to the control tags in the Ribbon XML file to identify the appropriate callback methods in your code.  

// For more information, see the Ribbon XML documentation in the Visual Studio Tools for Office Help.


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
            return Resources.image;
        }
        public void OnButtonClick(object sender)
        {
            ForwardSelectedEmail();
        }

        /*
        private void ForwardSelectedEmail()
        {
            // MessageBox.Show($"Input text: Winna");
            
            Outlook.Explorer explorer = Globals.ThisAddIn.Application.ActiveExplorer();

            if (explorer.Selection.Count > 0 && explorer.Selection[1] is Outlook.MailItem)
            {
                Outlook.MailItem selectedMail = explorer.Selection[1] as Outlook.MailItem;

                // Create a new mail item for forwarding
                Outlook.MailItem forwardMail = selectedMail.Forward();

                // Set the recipient's email address
                forwardMail.Recipients.Add("jay.truscott@outlook.com");

                // Optionally, modify other properties of the forwarded email
                // forwardMail.Subject = "Forwarded: " + selectedMail.Subject;

                // Send the forwarded email
                forwardMail.Send();
            }
            else
            {
                // Handle when no email is selected
                System.Windows.Forms.MessageBox.Show("Please select an email to forward.", "No Email Selected", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Warning);
            }
            
        }
        */

        public void ForwardSelectedEmail()
        {
            Outlook.Explorer explorer = Globals.ThisAddIn.Application.ActiveExplorer();

            if (explorer.Selection.Count == 0)
            {
                System.Windows.Forms.MessageBox.Show("Please select an email to forward.", "No Email Selected", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Warning);
            }
            else if (explorer.Selection.Count == 1 && explorer.Selection[1] is Outlook.MailItem)
            {
                string address = "jay.truscott@outlook.com";
                Outlook.MailItem selectedMail = explorer.Selection[1] as Outlook.MailItem;

                DialogResult result = MessageBox.Show($"Do you want to formard '{selectedMail.Subject}' to {address}?", "Confirmation", MessageBoxButtons.OKCancel);

                if (result == DialogResult.OK)
                {
                    Outlook.MailItem forwardMail = selectedMail.Forward();
                    forwardMail.Recipients.Add(address);
                    forwardMail.Subject = "Reported with Gone Phishing - " + selectedMail.Subject;
                    forwardMail.Send();
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
