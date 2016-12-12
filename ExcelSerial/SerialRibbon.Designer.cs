namespace ExcelSerial
{
    partial class SerialRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public SerialRibbon()
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
                readTask.Dispose();
                parseTask.Dispose();
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
            this.grpComSettings = this.Factory.CreateRibbonGroup();
            this.lblPort = this.Factory.CreateRibbonLabel();
            this.lblComParameter = this.Factory.CreateRibbonLabel();
            this.cmdRec = this.Factory.CreateRibbonButton();
            this.separator1 = this.Factory.CreateRibbonSeparator();
            this.lstPort = this.Factory.CreateRibbonDropDown();
            this.txtComSettings = this.Factory.CreateRibbonEditBox();
            this.cmdStop = this.Factory.CreateRibbonButton();
            this.grpData = this.Factory.CreateRibbonGroup();
            this.chkCsv = this.Factory.CreateRibbonCheckBox();
            this.chkFixLength = this.Factory.CreateRibbonCheckBox();
            this.chkBase64 = this.Factory.CreateRibbonCheckBox();
            this.separator2 = this.Factory.CreateRibbonSeparator();
            this.txtSeperator = this.Factory.CreateRibbonEditBox();
            this.txtLength = this.Factory.CreateRibbonEditBox();
            this.tab1.SuspendLayout();
            this.grpComSettings.SuspendLayout();
            this.grpData.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.grpComSettings);
            this.tab1.Groups.Add(this.grpData);
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // grpComSettings
            // 
            this.grpComSettings.Items.Add(this.lblPort);
            this.grpComSettings.Items.Add(this.lblComParameter);
            this.grpComSettings.Items.Add(this.cmdRec);
            this.grpComSettings.Items.Add(this.separator1);
            this.grpComSettings.Items.Add(this.lstPort);
            this.grpComSettings.Items.Add(this.txtComSettings);
            this.grpComSettings.Items.Add(this.cmdStop);
            this.grpComSettings.Label = "COM Einstellungen";
            this.grpComSettings.Name = "grpComSettings";
            // 
            // lblPort
            // 
            this.lblPort.Label = "Port";
            this.lblPort.Name = "lblPort";
            // 
            // lblComParameter
            // 
            this.lblComParameter.Label = "Parameter";
            this.lblComParameter.Name = "lblComParameter";
            // 
            // cmdRec
            // 
            this.cmdRec.Label = "REC";
            this.cmdRec.Name = "cmdRec";
            this.cmdRec.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.cmdRec_Click);
            // 
            // separator1
            // 
            this.separator1.Name = "separator1";
            // 
            // lstPort
            // 
            this.lstPort.Label = "dropDown1";
            this.lstPort.Name = "lstPort";
            this.lstPort.ShowLabel = false;
            // 
            // txtComSettings
            // 
            this.txtComSettings.Label = "Format";
            this.txtComSettings.MaxLength = 14;
            this.txtComSettings.Name = "txtComSettings";
            this.txtComSettings.ShowLabel = false;
            this.txtComSettings.Text = "460800/8-N-1";
            // 
            // cmdStop
            // 
            this.cmdStop.Label = "STOP";
            this.cmdStop.Name = "cmdStop";
            this.cmdStop.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.cmdStop_Click);
            // 
            // grpData
            // 
            this.grpData.Items.Add(this.chkCsv);
            this.grpData.Items.Add(this.chkFixLength);
            this.grpData.Items.Add(this.chkBase64);
            this.grpData.Items.Add(this.separator2);
            this.grpData.Items.Add(this.txtSeperator);
            this.grpData.Items.Add(this.txtLength);
            this.grpData.Label = "Datenformat";
            this.grpData.Name = "grpData";
            // 
            // chkCsv
            // 
            this.chkCsv.Label = "Seperator";
            this.chkCsv.Name = "chkCsv";
            this.chkCsv.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.chkCsv_Click);
            // 
            // chkFixLength
            // 
            this.chkFixLength.Label = "Fixe Länge";
            this.chkFixLength.Name = "chkFixLength";
            this.chkFixLength.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.chkFixLength_Click);
            // 
            // chkBase64
            // 
            this.chkBase64.Label = "Base64";
            this.chkBase64.Name = "chkBase64";
            // 
            // separator2
            // 
            this.separator2.Name = "separator2";
            // 
            // txtSeperator
            // 
            this.txtSeperator.Enabled = false;
            this.txtSeperator.Label = "editBox1";
            this.txtSeperator.MaxLength = 1;
            this.txtSeperator.Name = "txtSeperator";
            this.txtSeperator.ShowLabel = false;
            this.txtSeperator.Text = null;
            // 
            // txtLength
            // 
            this.txtLength.Enabled = false;
            this.txtLength.Label = "editBox1";
            this.txtLength.Name = "txtLength";
            this.txtLength.ShowLabel = false;
            this.txtLength.Text = null;
            // 
            // SerialRibbon
            // 
            this.Name = "SerialRibbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.SerialRibbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.grpComSettings.ResumeLayout(false);
            this.grpComSettings.PerformLayout();
            this.grpData.ResumeLayout(false);
            this.grpData.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpComSettings;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox txtComSettings;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel lblPort;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel lblComParameter;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator1;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown lstPort;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpData;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox chkCsv;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox chkFixLength;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator2;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox txtSeperator;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton cmdRec;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton cmdStop;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox chkBase64;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox txtLength;
    }

    partial class ThisRibbonCollection
    {
        internal SerialRibbon SerialRibbon
        {
            get { return this.GetRibbon<SerialRibbon>(); }
        }
    }
}
