using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Harmonix
{
    public partial class HarmonixRibbon
    {
        private void OpenFileBtn_Click(object sender, RibbonControlEventArgs e)
        {
            Harmonix.OpenCSV();
        }

        private void SettingsButton_Click(object sender, RibbonControlEventArgs e)
        {
            SettingsForm form = new SettingsForm();
            form.Show();
        }

        private void LockFirstColumnCheckBox_Click(object sender, RibbonControlEventArgs e)
        {
            RibbonCheckBox checkbox = (RibbonCheckBox)sender;

            Harmonix.GetActiveWindow().FreezePanes = false;

            if (checkbox.Checked)
            {
                Harmonix.GetActiveWindow().ScrollColumn = 1;
                Harmonix.GetActiveWindow().SplitColumn = 1;
            } else
            {
                Harmonix.GetActiveWindow().SplitColumn = 0;
            }

            Harmonix.GetActiveWindow().FreezePanes = true;
        }

        private void RunHarmonyButton_Click(object sender, RibbonControlEventArgs e)
        {
            string loadedFile = Harmonix.GetActiveWorkbook().FullName;

            if(!System.IO.File.Exists(loadedFile))
            {
                MessageBox.Show("Could not find the loaded CSV");
                return;
            }

            string harmonyPath = Settings.GetHarmonyPath();

            if(System.IO.File.Exists(harmonyPath))
            {
                Process.Start(harmonyPath, loadedFile);
            } else
            {
                MessageBox.Show("Could not find Harmony, please make sure the path is set correct.");
            }
        }
    }
}
