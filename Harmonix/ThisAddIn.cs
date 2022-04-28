using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Microsoft.Office.Interop.Excel;
using System.Diagnostics;
using System.IO;

namespace Harmonix
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            // Check if the Black Ops 3 root has been set yet
            if (Properties.Settings.Default.BlackOps3Root == string.Empty)
            {
                Settings.SetBlackOps3Root();
            }

            Application.WorkbookOpen += Application_WorkbookOpen;
        }

        private void Application_WorkbookOpen(Workbook Wb)
        {
            if(Settings.RunOnOpenCSV() && !Harmonix.OpenWithHarmonix)
            {
                string file = Wb.FullName;

                if (Path.GetExtension(file) == ".csv")
                {
                    Harmonix.SetupWorksheet();
                }
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

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
