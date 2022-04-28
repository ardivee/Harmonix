using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace Harmonix
{
    internal class Harmonix
    {
        /// <summary>
        /// If it's openened with "Harmonix" instead of via Excel open file
        /// </summary>
        public static bool OpenWithHarmonix = false;

        /// <summary>
        /// Open the Black Ops 3 CSV and setup the Worksheet
        /// </summary>
        public static void OpenCSV()
        {
            dynamic aliasFile = GetApp().GetOpenFilename(FileFilter: "Black Ops 3 Alias (*.csv), *.csv", Title: "Black Ops 3 Alias to Open");

            // Cancel if it's a bool
            if (aliasFile.GetType() == typeof(bool)) return;

            GetActiveWorkbook().Close();

            OpenWithHarmonix = true;

            GetApp().Workbooks.Open(aliasFile);

            SetupWorksheet();
        }

        /// <summary>
        /// Setup the worksheet with cool stuff
        /// </summary>
        public static void SetupWorksheet()
        {
            Worksheet worksheet = GetActiveWorksheet();

            int aliasRowCount = worksheet.UsedRange.Rows.Count;
            int aliasColumnCount = worksheet.UsedRange.Columns.Count;
            
            if(Settings.DarkThemeEnabled())
            {
                ActivateDarkMode(worksheet);
            }

            FreezeFirstRow();

            AutoFitCells(worksheet);

            SetToolTips(worksheet, aliasRowCount, aliasColumnCount);

            SetCellDataValidation(worksheet, aliasRowCount, aliasColumnCount);

            LockFirstRow(worksheet);

            GetActiveWorksheet().Change += Worksheet_Change;
            GetActiveWorksheet().BeforeDoubleClick += Worksheet_BeforeDoubleClick;

            OpenWithHarmonix = false;
        }

        /// <summary>
        /// Event handler before when the user double clicks a cell
        /// </summary>
        /// <param name="Target"></param>
        /// <param name="Cancel"></param>
        private static void Worksheet_BeforeDoubleClick(Range Target, ref bool Cancel)
        {
            // Get the header name
            string headerName = GetActiveWorksheet().Cells[1, Target.Column].Value;
            if (headerName == "FileSpecSustain" || headerName == "FileSpecRelease" || headerName == "FileSpec")
            {
                System.Windows.Forms.OpenFileDialog ofd = new System.Windows.Forms.OpenFileDialog();

                string value = Target.Value;

                string basePath = Path.Combine(Settings.GetBlackOps3Root(), "sound_assets");

                ofd.InitialDirectory = basePath;

                if (!string.IsNullOrWhiteSpace(value))
                {
                    string soundPath = Path.GetDirectoryName(Path.Combine(basePath, value));

                    if (Directory.Exists(soundPath))
                    {
                        ofd.InitialDirectory = soundPath;
                    }
                }

                ofd.Filter = "WAVE Files (.wav)|*.wav";
                ofd.Title = "Select a WAV File";

                if (ofd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    Target.Value = ofd.FileName.Replace(basePath + "\\", "");

                    Cancel = true;
                }
            }
        }

        /// <summary>
        /// Event handler when something changes in the worksheet
        /// </summary>
        /// <param name="Target"></param>
        private static void Worksheet_Change(Range Target)
        {
            Worksheet worksheet = GetActiveWorksheet();
            bool isProtected = worksheet.ProtectContents;

            if (isProtected)
            {
                worksheet.Unprotect();
            }

            var toolTip = BlackOps3Alias.GetToolTip(Target.Value);
            if (!string.IsNullOrWhiteSpace(toolTip))
            {
                Debug.WriteLine("Got a ToolTip for ya");
                Target.ClearComments();
                Comment comment = Target.AddComment(toolTip);
                comment.Shape.TextFrame.AutoSize = true;
            }

            BlackOps3Alias.DataValidationContext(Target);

            if (isProtected)
            {
                ProtectWorksheet(worksheet);
            }
        }

        /// <summary>
        /// Get the Application
        /// </summary>
        /// <returns></returns>
        public static Application GetApp()
        {
            return Globals.ThisAddIn.Application;
        }

        /// <summary>
        /// Get the active Workshop
        /// </summary>
        /// <returns></returns>
        public static Workbook GetActiveWorkbook()
        {
            return GetApp().ActiveWorkbook;
        }

        /// <summary>
        /// Get the active Worksheet
        /// </summary>
        /// <returns></returns>
        public static Worksheet GetActiveWorksheet()
        {
            return GetApp().ActiveWorkbook.ActiveSheet;
        }

        /// <summary>
        /// Get the active Window
        /// </summary>
        /// <returns></returns>
        public static Window GetActiveWindow()
        {
            return GetApp().ActiveWindow;
        }

        /// <summary>
        /// Set the status
        /// </summary>
        /// <param name="text"></param>
        public static void SetStatus(string text)
        {
            GetApp().StatusBar = text;
        }

        /// <summary>
        /// Hide the status bar
        /// </summary>
        public static void HideStatusBar()
        {
            GetApp().DisplayStatusBar = false;
        }

        /// <summary>
        /// Show the status bar
        /// </summary>
        public static void ShowStatusBar()
        {
            GetApp().DisplayStatusBar = true;
        }

        /// <summary>
        /// Resize all cells to auto fit the width
        /// </summary>
        /// <param name="worksheet"></param>
        public static void AutoFitCells(Worksheet worksheet)
        {
            bool isProtected = worksheet.ProtectContents;

            if (isProtected)
            {
                worksheet.Unprotect();
            }

            worksheet.Rows.AutoFit();
            worksheet.Columns.AutoFit();

            if (isProtected)
            {
                ProtectWorksheet(worksheet);
            }
        }

        /// <summary>
        /// Protect the worksheet
        /// </summary>
        /// <param name="worksheet"></param>
        public static void ProtectWorksheet(Worksheet worksheet)
        {
            worksheet.Protect(null, false, true, false, false, false, true, true, true, true, false, true, true, true, true, true);
        }

        /// <summary>
        /// This freezes the first row, so when we scroll down it's still visible
        /// </summary>
        public static void FreezeFirstRow()
        {
            GetActiveWindow().SplitRow = 1;
            GetActiveWindow().FreezePanes = true;
        }

        /// <summary>
        /// Lock the first row / header so we can't mess it up
        /// </summary>
        /// <param name="worksheet"></param>
        public static void LockFirstRow(Worksheet worksheet)
        {
            Range range = worksheet.Rows;
            range.Locked = false;

            // The first row / header
            Range header = worksheet.Rows[1];
            header.Locked = true;

            ProtectWorksheet(worksheet);

            // Enable selection on all other cells except the header
            worksheet.EnableSelection = XlEnableSelection.xlUnlockedCells;
        }

        /// <summary>
        /// Set a Dark Theme so our eyes don't get burned
        /// </summary>
        /// <param name="worksheet"></param>
        public static void ActivateDarkMode(Worksheet worksheet)
        {
            // All the cells
            Range range = worksheet.Rows;
            range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(66, 66, 66));
            range.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(250, 250, 250));
            range.Borders.LineStyle = XlLineStyle.xlContinuous;
            range.Borders.Weight = XlBorderWeight.xlThin;

            // The first row / header
            Range header = worksheet.Rows[1];
            header.Font.Bold = true;
            header.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(250, 250, 250));
            header.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(18, 18, 18));
        }

        /// <summary>
        /// Set ToolTips for cells that have a ToolTip
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="aliasRowCount"></param>
        /// <param name="aliasColumnCount"></param>
        public static void SetToolTips(Worksheet worksheet, int aliasRowCount, int aliasColumnCount)
        {
            for (int i = 1; i <= aliasRowCount; i++)
            {
                for (int j = 1; j <= aliasColumnCount; j++)
                {
                    Range cell = worksheet.Cells[i, j];

                    var tooltip = BlackOps3Alias.GetToolTip(cell.Value);
                    if (!string.IsNullOrWhiteSpace(tooltip))
                    {
                        Comment comment = cell.AddComment(tooltip);
                        comment.Shape.TextFrame.AutoSize = true;
                    }
                }
            }
        }

        /// <summary>
        /// Set Data Validation on cells that have any dropdown info
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="aliasRowCount"></param>
        /// <param name="aliasColumnCount"></param>
        public static void SetCellDataValidation(Worksheet worksheet, int aliasRowCount, int aliasColumnCount)
        {
            for (int i = 1; i <= aliasColumnCount; i++)
            {
                // Get the header cell
                Range cell = worksheet.Rows[1].Cells[i];

                // Get the entire column
                Range col = worksheet.Columns[i];

                // Set data validation if supported.
                BlackOps3Alias.CellDataValidation(cell.Value, col);
            }

            // Check for cells that need a data context value
            for (int i = 2; i <= aliasRowCount; i++)
            {
                for (int j = 1; j <= aliasColumnCount; j++)
                {
                    Range Target = worksheet.Cells[i, j];

                    BlackOps3Alias.DataValidationContext(Target, true);
                }
            }
        }
    }
}
