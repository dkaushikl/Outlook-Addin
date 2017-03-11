namespace OutlookMail
{
    partial class MailRead : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public MailRead()
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MailRead));
            this.tab1 = this.Factory.CreateRibbonTab();
            this.grpmailread = this.Factory.CreateRibbonGroup();
            this.ReadEmail = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.grpmailread.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.grpmailread);
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // grpmailread
            // 
            this.grpmailread.Items.Add(this.ReadEmail);
            this.grpmailread.Label = "Mail Read";
            this.grpmailread.Name = "grpmailread";
            // 
            // ReadEmail
            // 
            this.ReadEmail.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.ReadEmail.Description = "Message after email sending";
            this.ReadEmail.Image = ((System.Drawing.Image)(resources.GetObject("ReadEmail.Image")));
            this.ReadEmail.ImageName = "ReadEmail";
            this.ReadEmail.KeyTip = "Y";
            this.ReadEmail.Label = "ReadEmail";
            this.ReadEmail.Name = "ReadEmail";
            this.ReadEmail.ShowImage = true;
            this.ReadEmail.SuperTip = "ReadEmail";
            this.ReadEmail.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ReadEmail_Click);
            // 
            // MailRead
            // 
            this.Name = "MailRead";
            this.RibbonType = "Microsoft.Outlook.Explorer";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.MailRead_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.grpmailread.ResumeLayout(false);
            this.grpmailread.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpmailread;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ReadEmail;
    }

    partial class ThisRibbonCollection
    {
        internal MailRead MailRead
        {
            get { return this.GetRibbon<MailRead>(); }
        }
    }
}
