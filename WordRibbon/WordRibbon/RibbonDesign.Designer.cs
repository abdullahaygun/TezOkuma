
namespace WordRibbon
{
    partial class RibbonDesign : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public RibbonDesign()
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
            this.item1 = this.Factory.CreateRibbonGroup();
            this.btn = this.Factory.CreateRibbonButton();
            this.button1 = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.item1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.ControlId.OfficeId = "TabHome";
            this.tab1.Groups.Add(this.item1);
            this.tab1.Label = "TabHome";
            this.tab1.Name = "tab1";
            // 
            // item1
            // 
            this.item1.Items.Add(this.btn);
            this.item1.Label = "Click";
            this.item1.Name = "item1";
            // 
            // btn
            // 
            this.btn.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btn.Label = "Tıkla";
            this.btn.Name = "btn";
            this.btn.ShowImage = true;
            this.btn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_Click);
            // 
            // button1
            // 
            this.button1.Label = "deneme";
            this.button1.Name = "button1";
            this.button1.ShowImage = true;
            // 
            // RibbonDesign
            // 
            this.Name = "RibbonDesign";
            // 
            // RibbonDesign.OfficeMenu
            // 
            this.OfficeMenu.Items.Add(this.button1);
            this.RibbonType = "Microsoft.Word.Document";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.RibbonDesign_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.item1.ResumeLayout(false);
            this.item1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup item1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button1;
    }

    partial class ThisRibbonCollection
    {
        internal RibbonDesign RibbonDesign
        {
            get { return this.GetRibbon<RibbonDesign>(); }
        }
    }
}
