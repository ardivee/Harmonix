using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Harmonix
{
    internal class Settings
    {
        /// <summary>
        /// Get the Black Ops 3 Root folder
        /// </summary>
        /// <returns></returns>
        public static string GetBlackOps3Root()
        {
            return Properties.Settings.Default.BlackOps3Root;
        }

        /// <summary>
        /// Get the Harmony Path
        /// </summary>
        /// <returns></returns>
        public static string GetHarmonyPath()
        {
            return Properties.Settings.Default.HarmonyPath;
        }

        /// <summary>
        /// Gets if we should run on opening a CSV file
        /// </summary>
        /// <returns></returns>
        public static bool RunOnOpenCSV()
        {
            return Properties.Settings.Default.RunOnOpenCSV;
        }

        /// <summary>
        /// Gets if the Dark Theme is enabled
        /// </summary>
        /// <returns></returns>
        public static bool DarkThemeEnabled()
        {
            return Properties.Settings.Default.DarkTheme;
        }

        /// <summary>
        /// Set the Black Ops 3 root folder
        /// </summary>
        public static void SetBlackOps3Root()
        {
            System.Windows.Forms.DialogResult result = System.Windows.Forms.MessageBox.Show($"Black Ops 3 Root has not been set yet. {Environment.NewLine}Do you want to auto search for it ?", "Black Ops 3 Root", System.Windows.Forms.MessageBoxButtons.YesNo);

            if (result == System.Windows.Forms.DialogResult.Yes)
            {
                string bo3Root = Environment.GetEnvironmentVariable("TA_TOOLS_PATH");

                if (!string.IsNullOrWhiteSpace(bo3Root))
                {
                    Properties.Settings.Default.BlackOps3Root = bo3Root;

                    System.Windows.Forms.MessageBox.Show($"Black Ops 3 Root has been set {Environment.NewLine}{bo3Root}");
                }
                else
                {
                    System.Windows.Forms.MessageBox.Show("Could not find it, please manually select the root.");

                    var dialog = new FolderSelectDialog
                    {
                        InitialDirectory = @"C:\",
                        Title = "Select your Black Ops 3 Root Folder"
                    };

                    if (dialog.Show())
                    {
                        if (!string.IsNullOrWhiteSpace(dialog.FileName))
                        {
                            Properties.Settings.Default.BlackOps3Root = dialog.FileName;
                        }
                    }
                }
            }
            else
            {
                var dialog = new FolderSelectDialog
                {
                    InitialDirectory = @"C:\",
                    Title = "Select your Black Ops 3 Root Folder"
                };

                if (dialog.Show())
                {
                    if (!string.IsNullOrWhiteSpace(dialog.FileName))
                    {
                        Properties.Settings.Default.BlackOps3Root = dialog.FileName;

                        System.Windows.Forms.MessageBox.Show($"Black Ops 3 Root has been set {Environment.NewLine}{dialog.FileName}");
                    }
                }
            }

            Properties.Settings.Default.Save();
        }
    }
}
