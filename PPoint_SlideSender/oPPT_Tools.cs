using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Runtime.InteropServices;

namespace PPoint_SlideSender
{
    public partial class SlideSender
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        /*Microsoft.Win32.RegistryKey key;
        key = Microsoft.Win32.Registry.LocalMachine.CreateSubKey("SOFTWARE\\MICROSOFT\\.NETFramework\\Security\\TrustManager\\PromptingLevel");
        key.SetValue("MyComputer", "AuthenticodeRequired");
        //key.SetValue("LocalIntranet", "AuthenticodeRequired");
        //key.SetValue("Internet", "AuthenticodeRequired");
        //key.SetValue("TrustedSites", "AuthenticodeRequired");
        //key.SetValue("UntrustedSites", "Disabled");
        key.Close();
         */

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
