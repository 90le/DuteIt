namespace DuteIT
{
    partial class DuteIT : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public DuteIT()
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(DuteIT));
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group5 = this.Factory.CreateRibbonGroup();
            this.menu23 = this.Factory.CreateRibbonMenu();
            this.button127 = this.Factory.CreateRibbonButton();
            this.button79 = this.Factory.CreateRibbonButton();
            this.button128 = this.Factory.CreateRibbonButton();
            this.group3 = this.Factory.CreateRibbonGroup();
            this.box1 = this.Factory.CreateRibbonBox();
            this.splitButton1 = this.Factory.CreateRibbonSplitButton();
            this.button18 = this.Factory.CreateRibbonButton();
            this.menu2 = this.Factory.CreateRibbonMenu();
            this.button8 = this.Factory.CreateRibbonButton();
            this.button9 = this.Factory.CreateRibbonButton();
            this.button294 = this.Factory.CreateRibbonButton();
            this.button295 = this.Factory.CreateRibbonButton();
            this.button296 = this.Factory.CreateRibbonButton();
            this.button297 = this.Factory.CreateRibbonButton();
            this.menu4 = this.Factory.CreateRibbonMenu();
            this.menu1 = this.Factory.CreateRibbonMenu();
            this.button2 = this.Factory.CreateRibbonButton();
            this.button3 = this.Factory.CreateRibbonButton();
            this.button5 = this.Factory.CreateRibbonButton();
            this.button15 = this.Factory.CreateRibbonButton();
            this.button169 = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group5.SuspendLayout();
            this.group3.SuspendLayout();
            this.box1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.Groups.Add(this.group5);
            this.tab1.Groups.Add(this.group3);
            this.tab1.Label = "DTI Tool";
            this.tab1.Name = "tab1";
            // 
            // group5
            // 
            this.group5.Items.Add(this.menu23);
            this.group5.Items.Add(this.button127);
            this.group5.Items.Add(this.button79);
            this.group5.Items.Add(this.button128);
            this.group5.Label = "帮 助";
            this.group5.Name = "group5";
            // 
            // menu23
            // 
            this.menu23.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.menu23.Image = global::DuteIT.Properties.Resources.DTI;
            this.menu23.Label = "教程  ";
            this.menu23.Name = "menu23";
            this.menu23.ScreenTip = "教程";
            this.menu23.ShowImage = true;
            this.menu23.SuperTip = "DTI相关的教程链接";
            // 
            // button127
            // 
            this.button127.Enabled = false;
            this.button127.Label = " 设置 ";
            this.button127.Name = "button127";
            this.button127.ScreenTip = "DTI设置";
            this.button127.SuperTip = "设置需要账号密码，用于特别功能中。";
            this.button127.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button127_Click);
            // 
            // button79
            // 
            this.button79.Enabled = false;
            this.button79.Label = " 官网";
            this.button79.Name = "button79";
            this.button79.ScreenTip = "DTI官网";
            this.button79.SuperTip = "点击进入官网，下载最新版DTI插件、使用教程及更多资源";
            this.button79.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button79_Click_1);
            // 
            // button128
            // 
            this.button128.Label = " 关注";
            this.button128.Name = "button128";
            this.button128.ScreenTip = "关注我们";
            this.button128.SuperTip = "欢迎交流与建议";
            this.button128.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button128_Click_1);
            // 
            // group3
            // 
            this.group3.Items.Add(this.box1);
            this.group3.Items.Add(this.menu2);
            this.group3.Items.Add(this.menu4);
            this.group3.Label = "常用功能";
            this.group3.Name = "group3";
            // 
            // box1
            // 
            this.box1.BoxStyle = Microsoft.Office.Tools.Ribbon.RibbonBoxStyle.Vertical;
            this.box1.Items.Add(this.splitButton1);
            this.box1.Name = "box1";
            // 
            // splitButton1
            // 
            this.splitButton1.Checked = true;
            this.splitButton1.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.splitButton1.Image = global::DuteIT.Properties.Resources.AllBorders;
            this.splitButton1.Items.Add(this.button18);
            this.splitButton1.Label = "合并相同";
            this.splitButton1.Name = "splitButton1";
            this.splitButton1.ScreenTip = "合并单元格";
            this.splitButton1.SuperTip = "合并相同内容的单元格";
            this.splitButton1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.splitButton1_Click);
            // 
            // button18
            // 
            this.button18.Image = ((System.Drawing.Image)(resources.GetObject("button18.Image")));
            this.button18.Label = "拆分单元格";
            this.button18.Name = "button18";
            this.button18.ScreenTip = "拆分单元格";
            this.button18.ShowImage = true;
            this.button18.SuperTip = "拆分合并后的单元格";
            this.button18.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button18_Click);
            // 
            // menu2
            // 
            this.menu2.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.menu2.Image = global::DuteIT.Properties.Resources.Find;
            this.menu2.Items.Add(this.button8);
            this.menu2.Items.Add(this.button9);
            this.menu2.Items.Add(this.button294);
            this.menu2.Items.Add(this.button295);
            this.menu2.Items.Add(this.button296);
            this.menu2.Items.Add(this.button297);
            this.menu2.Label = "数据定位";
            this.menu2.Name = "menu2";
            this.menu2.ShowImage = true;
            // 
            // button8
            // 
            this.button8.Image = global::DuteIT.Properties.Resources.SelEmptyCells;
            this.button8.Label = "定位批注";
            this.button8.Name = "button8";
            this.button8.ScreenTip = "定位批注";
            this.button8.ShowImage = true;
            this.button8.SuperTip = "定位选中区域包含批注的单元格";
            // 
            // button9
            // 
            this.button9.Image = global::DuteIT.Properties.Resources.SelEmptyCells;
            this.button9.Label = "定位空值";
            this.button9.Name = "button9";
            this.button9.ScreenTip = "定位空值";
            this.button9.ShowImage = true;
            this.button9.SuperTip = "定位选中区内容为空的单元格";
            // 
            // button294
            // 
            this.button294.Image = global::DuteIT.Properties.Resources.SelEmptyCells;
            this.button294.Label = "定位文本";
            this.button294.Name = "button294";
            this.button294.ScreenTip = "定位文本";
            this.button294.ShowImage = true;
            this.button294.SuperTip = "定位选中区域的值为文本类型的单元格";
            // 
            // button295
            // 
            this.button295.Image = global::DuteIT.Properties.Resources.SelEmptyCells;
            this.button295.Label = "定位对象";
            this.button295.Name = "button295";
            this.button295.ScreenTip = "定位对象";
            this.button295.ShowImage = true;
            this.button295.SuperTip = "定位当前工作表悬浮的窗体对象";
            // 
            // button296
            // 
            this.button296.Image = global::DuteIT.Properties.Resources.SelEmptyCells;
            this.button296.Label = "定位数值";
            this.button296.Name = "button296";
            this.button296.ScreenTip = "定位数值";
            this.button296.ShowImage = true;
            this.button296.SuperTip = "定位选中区内容属于数字格式的单元格";
            // 
            // button297
            // 
            this.button297.Image = global::DuteIT.Properties.Resources.SelEmptyCells;
            this.button297.Label = "定位错误";
            this.button297.Name = "button297";
            this.button297.ScreenTip = "定位错误";
            this.button297.ShowImage = true;
            this.button297.SuperTip = "定位选中区域存在公式错误的单元格";
            // 
            // menu4
            // 
            this.menu4.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.menu4.Image = global::DuteIT.Properties.Resources.folder_add;
            this.menu4.Items.Add(this.menu1);
            this.menu4.Items.Add(this.button5);
            this.menu4.Items.Add(this.button15);
            this.menu4.Items.Add(this.button169);
            this.menu4.Label = "文本处理";
            this.menu4.Name = "menu4";
            this.menu4.ShowImage = true;
            // 
            // menu1
            // 
            this.menu1.Image = global::DuteIT.Properties.Resources.repeat_records;
            this.menu1.Items.Add(this.button2);
            this.menu1.Items.Add(this.button3);
            this.menu1.Label = "关键词处理";
            this.menu1.Name = "menu1";
            this.menu1.ShowImage = true;
            // 
            // button2
            // 
            this.button2.Image = global::DuteIT.Properties.Resources.OnlyLeftAndRightBorders;
            this.button2.Label = "关键词上色";
            this.button2.Name = "button2";
            this.button2.ScreenTip = "关键词上色";
            this.button2.ShowImage = true;
            this.button2.SuperTip = "找出单元格内指定的关键词，并设置颜色。";
            this.button2.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button2_Click);
            // 
            // button3
            // 
            this.button3.Image = global::DuteIT.Properties.Resources.OnlyLeftAndRightBorders;
            this.button3.Label = "提取关键词";
            this.button3.Name = "button3";
            this.button3.ScreenTip = "提取关键词";
            this.button3.ShowImage = true;
            this.button3.SuperTip = "查找内容中是否包含相应的关键词，并提取出来。";
            this.button3.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button3_Click);
            // 
            // button5
            // 
            this.button5.Image = global::DuteIT.Properties.Resources.Find;
            this.button5.Label = "字符串加密";
            this.button5.Name = "button5";
            this.button5.ScreenTip = "字符串加密";
            this.button5.ShowImage = true;
            this.button5.SuperTip = "对选中区域所有内容使用SHA256加密";
            this.button5.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button5_Click);
            // 
            // button15
            // 
            this.button15.Image = global::DuteIT.Properties.Resources.application_form_edit;
            this.button15.Label = "匹配删除行";
            this.button15.Name = "button15";
            this.button15.ScreenTip = "匹配删除行";
            this.button15.ShowImage = true;
            this.button15.SuperTip = "对选中区域的内容做正则表达式的匹配，命中的字符行会被进行删除。";
            this.button15.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button15_Click);
            // 
            // button169
            // 
            this.button169.Image = global::DuteIT.Properties.Resources.script_code;
            this.button169.Label = "公式转数值";
            this.button169.Name = "button169";
            this.button169.ScreenTip = "转换数值";
            this.button169.ShowImage = true;
            this.button169.SuperTip = "对选中区域或当前工作表的所有内容去除公式";
            this.button169.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button169_Click_1);
            // 
            // DuteIT
            // 
            this.Name = "DuteIT";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.DuteIT_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group5.ResumeLayout(false);
            this.group5.PerformLayout();
            this.group3.ResumeLayout(false);
            this.group3.PerformLayout();
            this.box1.ResumeLayout(false);
            this.box1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group5;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menu23;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button127;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button79;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button128;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group3;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button169;
        private Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button2;
        internal Microsoft.Office.Tools.Ribbon.RibbonSplitButton splitButton1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button18;
        internal Microsoft.Office.Tools.Ribbon.RibbonBox box1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button3;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button5;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menu1;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menu2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button8;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button9;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button294;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button295;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button296;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button297;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button15;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menu4;
    }

    partial class ThisRibbonCollection
    {
        public object CustomRibbon { get; internal set; }

        internal DuteIT DuteIT
        {
            get { return this.GetRibbon<DuteIT>(); }
        }
    }
}
