namespace PowerPointSetAudio
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
            this.tab1 = this.Factory.CreateRibbonTab();
            this.AudioSetting = this.Factory.CreateRibbonGroup();
            this.buttonClearAudio = this.Factory.CreateRibbonButton();
            this.buttonSetAudio = this.Factory.CreateRibbonButton();
            this.buttonExportMP4 = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.AudioSetting.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.AudioSetting);
            this.tab1.Label = "PPTX2MP4";
            this.tab1.Name = "tab1";
            // 
            // AudioSetting
            // 
            this.AudioSetting.Items.Add(this.buttonClearAudio);
            this.AudioSetting.Items.Add(this.buttonSetAudio);
            this.AudioSetting.Items.Add(this.buttonExportMP4);
            this.AudioSetting.Label = "AudioSetting";
            this.AudioSetting.Name = "AudioSetting";
            // 
            // buttonClearAudio
            // 
            this.buttonClearAudio.Image = global::PowerPointSetAudio.Properties.Resources.OIP_C;
            this.buttonClearAudio.Label = "Clear All Audio";
            this.buttonClearAudio.Name = "buttonClearAudio";
            this.buttonClearAudio.ShowImage = true;
            this.buttonClearAudio.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonClearAudio_Click);
            // 
            // buttonSetAudio
            // 
            this.buttonSetAudio.Image = global::PowerPointSetAudio.Properties.Resources.tr;
            this.buttonSetAudio.Label = "Set All Audio";
            this.buttonSetAudio.Name = "buttonSetAudio";
            this.buttonSetAudio.ShowImage = true;
            this.buttonSetAudio.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button1_Click);
            // 
            // buttonExportMP4
            // 
            this.buttonExportMP4.Image = global::PowerPointSetAudio.Properties.Resources.f_;
            this.buttonExportMP4.Label = "Export MP4";
            this.buttonExportMP4.Name = "buttonExportMP4";
            this.buttonExportMP4.ShowImage = true;
            this.buttonExportMP4.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonExportMP4_Click);
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.PowerPoint.Presentation";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.AudioSetting.ResumeLayout(false);
            this.AudioSetting.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup AudioSetting;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonSetAudio;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonExportMP4;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonClearAudio;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
