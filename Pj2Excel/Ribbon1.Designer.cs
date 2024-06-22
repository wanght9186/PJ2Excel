namespace Pj2Excel
{
    partial class Ribbon1 : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon1()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region 组件设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.Pj2ExcelTab = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.ToExcel = this.Factory.CreateRibbonButton();
            this.Pj2ExcelTab.SuspendLayout();
            this.group1.SuspendLayout();
            this.SuspendLayout();
            // 
            // Pj2ExcelTab
            // 
            this.Pj2ExcelTab.Groups.Add(this.group1);
            this.Pj2ExcelTab.Label = "Pj2Excel";
            this.Pj2ExcelTab.Name = "Pj2ExcelTab";
            // 
            // group1
            // 
            this.group1.Items.Add(this.ToExcel);
            this.group1.Name = "group1";
            // 
            // ToExcel
            // 
            this.ToExcel.Label = "生成表";
            this.ToExcel.Name = "ToExcel";
            this.ToExcel.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ToExcel_Click);
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.Project.Project";
            this.Tabs.Add(this.Pj2ExcelTab);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.Pj2ExcelTab.ResumeLayout(false);
            this.Pj2ExcelTab.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab Pj2ExcelTab;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ToExcel;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
