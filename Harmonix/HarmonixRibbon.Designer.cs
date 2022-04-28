namespace Harmonix
{
    partial class HarmonixRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public HarmonixRibbon()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

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

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.OpenFileBtn = this.Factory.CreateRibbonButton();
            this.SettingsButton = this.Factory.CreateRibbonButton();
            this.LockFirstColumnCheckBox = this.Factory.CreateRibbonCheckBox();
            this.RunHarmonyButton = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group1);
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.OpenFileBtn);
            this.group1.Items.Add(this.SettingsButton);
            this.group1.Items.Add(this.RunHarmonyButton);
            this.group1.Items.Add(this.LockFirstColumnCheckBox);
            this.group1.Label = "Harmonix";
            this.group1.Name = "group1";
            // 
            // OpenFileBtn
            // 
            this.OpenFileBtn.Label = "Open File";
            this.OpenFileBtn.Name = "OpenFileBtn";
            this.OpenFileBtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.OpenFileBtn_Click);
            // 
            // SettingsButton
            // 
            this.SettingsButton.Label = "Settings";
            this.SettingsButton.Name = "SettingsButton";
            this.SettingsButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.SettingsButton_Click);
            // 
            // LockFirstColumnCheckBox
            // 
            this.LockFirstColumnCheckBox.Label = "Lock First Column";
            this.LockFirstColumnCheckBox.Name = "LockFirstColumnCheckBox";
            this.LockFirstColumnCheckBox.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.LockFirstColumnCheckBox_Click);
            // 
            // RunHarmonyButton
            // 
            this.RunHarmonyButton.Label = "Run Harmony";
            this.RunHarmonyButton.Name = "RunHarmonyButton";
            this.RunHarmonyButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.RunHarmonyButton_Click);
            // 
            // HarmonixRibbon
            // 
            this.Name = "HarmonixRibbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton OpenFileBtn;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton SettingsButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox LockFirstColumnCheckBox;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton RunHarmonyButton;
    }

    partial class ThisRibbonCollection
    {
        internal HarmonixRibbon HarmonixRibbon
        {
            get { return this.GetRibbon<HarmonixRibbon>(); }
        }
    }
}
