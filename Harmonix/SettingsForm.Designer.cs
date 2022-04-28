namespace Harmonix
{
    partial class SettingsForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(SettingsForm));
            this.DarkThemeCheckbox = new System.Windows.Forms.CheckBox();
            this.SetBlackOps3RootButton = new System.Windows.Forms.Button();
            this.RunOnOpenCSVCheckBox = new System.Windows.Forms.CheckBox();
            this.LocateHarmonyButton = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // DarkThemeCheckbox
            // 
            this.DarkThemeCheckbox.AutoSize = true;
            this.DarkThemeCheckbox.Location = new System.Drawing.Point(12, 70);
            this.DarkThemeCheckbox.Name = "DarkThemeCheckbox";
            this.DarkThemeCheckbox.Size = new System.Drawing.Size(85, 17);
            this.DarkThemeCheckbox.TabIndex = 0;
            this.DarkThemeCheckbox.Text = "Dark Theme";
            this.DarkThemeCheckbox.UseVisualStyleBackColor = true;
            this.DarkThemeCheckbox.Click += new System.EventHandler(this.DarkThemeCheckbox_Click);
            // 
            // SetBlackOps3RootButton
            // 
            this.SetBlackOps3RootButton.Location = new System.Drawing.Point(12, 12);
            this.SetBlackOps3RootButton.Name = "SetBlackOps3RootButton";
            this.SetBlackOps3RootButton.Size = new System.Drawing.Size(136, 23);
            this.SetBlackOps3RootButton.TabIndex = 1;
            this.SetBlackOps3RootButton.Text = "Set Black Ops 3 Root";
            this.SetBlackOps3RootButton.UseVisualStyleBackColor = true;
            this.SetBlackOps3RootButton.Click += new System.EventHandler(this.SetBlackOps3RootButton_Click);
            // 
            // RunOnOpenCSVCheckBox
            // 
            this.RunOnOpenCSVCheckBox.AutoSize = true;
            this.RunOnOpenCSVCheckBox.Location = new System.Drawing.Point(12, 93);
            this.RunOnOpenCSVCheckBox.Name = "RunOnOpenCSVCheckBox";
            this.RunOnOpenCSVCheckBox.Size = new System.Drawing.Size(114, 17);
            this.RunOnOpenCSVCheckBox.TabIndex = 2;
            this.RunOnOpenCSVCheckBox.Text = "Run on Open CSV";
            this.RunOnOpenCSVCheckBox.UseVisualStyleBackColor = true;
            this.RunOnOpenCSVCheckBox.Click += new System.EventHandler(this.RunOnOpenCSVCheckBox_Click);
            // 
            // LocateHarmonyButton
            // 
            this.LocateHarmonyButton.Location = new System.Drawing.Point(12, 41);
            this.LocateHarmonyButton.Name = "LocateHarmonyButton";
            this.LocateHarmonyButton.Size = new System.Drawing.Size(136, 23);
            this.LocateHarmonyButton.TabIndex = 3;
            this.LocateHarmonyButton.Text = "Locate Harmony";
            this.LocateHarmonyButton.UseVisualStyleBackColor = true;
            this.LocateHarmonyButton.Click += new System.EventHandler(this.LocateHarmonyButton_Click);
            // 
            // SettingsForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(288, 153);
            this.Controls.Add(this.LocateHarmonyButton);
            this.Controls.Add(this.RunOnOpenCSVCheckBox);
            this.Controls.Add(this.SetBlackOps3RootButton);
            this.Controls.Add(this.DarkThemeCheckbox);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "SettingsForm";
            this.Text = "Harmonix - Settings";
            this.Load += new System.EventHandler(this.SettingsForm_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.CheckBox DarkThemeCheckbox;
        private System.Windows.Forms.Button SetBlackOps3RootButton;
        private System.Windows.Forms.CheckBox RunOnOpenCSVCheckBox;
        private System.Windows.Forms.Button LocateHarmonyButton;
    }
}