using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Harmonix
{
    public partial class SettingsForm : Form
    {
        public SettingsForm()
        {
            InitializeComponent();
        }

        private void SettingsForm_Load(object sender, EventArgs e)
        {
            DarkThemeCheckbox.Checked = Settings.DarkThemeEnabled();
            RunOnOpenCSVCheckBox.Checked = Settings.RunOnOpenCSV();
        }

        private void SetBlackOps3RootButton_Click(object sender, EventArgs e)
        {
            Settings.SetBlackOps3Root();
        }

        private void DarkThemeCheckbox_Click(object sender, EventArgs e)
        {
            CheckBox checkBox = (CheckBox)sender;

            Properties.Settings.Default.DarkTheme = checkBox.Checked;
            Properties.Settings.Default.Save();
        }

        private void RunOnOpenCSVCheckBox_Click(object sender, EventArgs e)
        {
            CheckBox checkBox = (CheckBox)sender;

            Properties.Settings.Default.RunOnOpenCSV = checkBox.Checked;
            Properties.Settings.Default.Save();
        }

        private void LocateHarmonyButton_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();

            ofd.Filter = "Harmony |Harmony.exe";
            ofd.Title = "Select Harmony";

            if (ofd.ShowDialog() == DialogResult.OK)
            {
                Properties.Settings.Default.HarmonyPath = ofd.FileName;
                Properties.Settings.Default.Save();
                MessageBox.Show("Harmony path have been saved");
            }
        }
    }
}
