using System;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;
using System.Windows.Forms;
using System.Diagnostics;
using Microsoft.Office.Interop.Excel;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Runtime.InteropServices;


namespace DuteIT
{
    public partial class DuteIT
    {
        public class Win32API
        {
            [DllImport("user32.dll", SetLastError = true)]
            public static extern IntPtr SetParent(IntPtr hWndChild, IntPtr hWndNewParent);
        }
        private void DuteIT_Load(object sender, RibbonUIEventArgs e)
        {
            // 激活指定选项卡
            tabActivate(tab1,"tab1");
        }

        /// <summary>
        /// 激活指定选项卡
        /// </summary>
        private static void  tabActivate(RibbonTab tab,string name) 
        {
            //RibbonTab label = (RibbonTab)tab;
            tab.RibbonUI.ActivateTab(name);
        }


        #region 注册自定义函数
        private void button30_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.Application.MacroOptions2(
                Macro: "test",
                Description: "test的函数描述说明",
                Category: "DTI函数",
                ArgumentDescriptions: new[] { "1", "2" }
            );
        }
        #endregion

        #region 数据定位 - 开始

        #region 数据定位：定位批注
        private void button8_Click_1(object sender, RibbonControlEventArgs e)
        {
            Excel.Application App = Globals.ThisAddIn.Application;
            Excel.Range Rng = App.Selection;//获取当前选中单元格
            Excel.Worksheet Sh = App.ActiveSheet;//获取当前选中工作表

            try
            {
                if (Rng.Count > 1)
                {
                    Rng.SpecialCells(Excel.XlCellType.xlCellTypeComments).Select();
                    //也可以写成这样Globals.ThisAddIn.Application.Selection.SpecialCells(2, 1).Select();
                }
                else
                {
                    Sh.Cells.SpecialCells(Excel.XlCellType.xlCellTypeComments).Select();
                }
            }
            catch (Exception)
            {
                MessageBox.Show("没有找到批注单元格。", "Excel", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
        #endregion

        #region 数据定位：定位数字
        private void button296_Click_1(object sender, RibbonControlEventArgs e)
        {
            Excel.Application App = Globals.ThisAddIn.Application;
            Excel.Range Rng = App.Selection;//获取当前选中单元格
            Excel.Worksheet Sh = App.ActiveSheet;//获取当前选中工作表

            try
            {

                if (Rng.Count > 1)
                {
                    Rng.SpecialCells(Excel.XlCellType.xlCellTypeConstants, 1).Select();
                }
                else
                {
                    Sh.Cells.SpecialCells(Excel.XlCellType.xlCellTypeConstants, 1).Select();
                }
            }
            catch (Exception)
            {
                MessageBox.Show("没有找到数字单元格。", "Excel", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
        #endregion

        #region 数据定位：定位文本
        private void button294_Click_1(object sender, RibbonControlEventArgs e)
        {
            Excel.Application App = Globals.ThisAddIn.Application;
            Excel.Range Rng = App.Selection;//获取当前选中单元格
            Excel.Worksheet Sh = App.ActiveSheet;//获取当前选中工作表

            try
            {

                if (Rng.Count > 1)
                {
                    Rng.SpecialCells(Excel.XlCellType.xlCellTypeConstants, 2).Select();
                }
                else
                {
                    Sh.Cells.SpecialCells(Excel.XlCellType.xlCellTypeConstants, 2).Select();
                }
            }
            catch (Exception)
            {
                MessageBox.Show("没有找到文本单元格。", "Excel", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
        #endregion

        #region 数据定位：定位空值
        private void button9_Click_1(object sender, RibbonControlEventArgs e)
        {
            Excel.Application App = Globals.ThisAddIn.Application;
            Excel.Range Rng = App.Selection;//获取当前选中单元格
            Excel.Worksheet Sh = App.ActiveSheet;//获取当前选中工作表

            try
            {

                if (Rng.Count > 1)
                {
                    Rng.SpecialCells(Excel.XlCellType.xlCellTypeBlanks).Select();
                }
                else
                {
                    Sh.Cells.SpecialCells(Excel.XlCellType.xlCellTypeBlanks).Select();
                }
            }
            catch (Exception)
            {
                MessageBox.Show("没有找到空单元格。", "Excel", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
        #endregion

        #region 数据定位：定位对象
        private void button295_Click_1(object sender, RibbonControlEventArgs e)
        {
            Excel.Application App = Globals.ThisAddIn.Application;
            Excel.Range Rng = App.Selection;//获取当前选中单元格
            Excel.Worksheet Sh = App.ActiveSheet;//获取当前选中工作表

            try
            {
                Globals.ThisAddIn.Application.ActiveSheet.DrawingObjects.Select();
            }
            catch (Exception)
            {
                MessageBox.Show("没有找到窗体对象。", "Excel", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
        #endregion

        #region 数据定位：定位错误
        private void button297_Click_1(object sender, RibbonControlEventArgs e)
        {
            Excel.Application App = Globals.ThisAddIn.Application;
            Excel.Range Rng = App.Selection;//获取当前选中单元格
            Excel.Worksheet Sh = App.ActiveSheet;//获取当前选中工作表

            try
            {

                if (Rng.Count > 1)
                {
                    Rng.SpecialCells(Excel.XlCellType.xlCellTypeFormulas, 16).Select();
                    //也可以写成这样Globals.ThisAddIn.Application.Selection.SpecialCells(2, 1).Select();
                }
                else
                {
                    Sh.Cells.SpecialCells(Excel.XlCellType.xlCellTypeFormulas, 16).Select();
                }
            }
            catch (Exception)
            {
                MessageBox.Show("没有找到错误单元格。", "Excel", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
        #endregion

        #endregion 数据定位 - 结束

        #region 数据转换：公式转数值
        private void button169_Click_1(object sender, RibbonControlEventArgs e)
        {
            Excel.Application App = Globals.ThisAddIn.Application;
            Excel.Range Rng = App.Selection;//获取当前选中单元格
            Excel.Worksheet Sh = App.ActiveSheet;//获取当前选中工作表

            try
            {
                if (Rng.Count > 1)
                {
                    DialogResult result = MessageBox.Show("确定对当前选定区域进行转数值操作吗？此操作不可撤销！", "数值转换", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);
                    if (result == DialogResult.OK)
                    {
                        Rng.Value = Rng.Value;
                    }
                    Rng.Select();
                }
                else
                {
                    DialogResult result = MessageBox.Show("确定对当前工作表进行转数值操作吗？此操作不可撤销！", "数值转换", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);
                    if (result == DialogResult.OK)
                    {
                        //取消筛选和隐藏，防止转数值时出错。
                        Sh.AutoFilterMode = false;
                        Sh.Cells.EntireColumn.Hidden = false;
                        Sh.Cells.EntireRow.Hidden = false;
                        //复制粘贴为值
                        Sh.Cells.Copy();
                        Sh.Cells[1.1].PasteSpecial(Excel.XlPasteType.xlPasteValues);
                    }
                    Rng.Select();
                }
            }
            catch (Exception)
            {
                MessageBox.Show("转换失败！！！", "Excel", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
        #endregion

        #region 数据转换：拆分单元格
        private void button18_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Application App = Globals.ThisAddIn.Application;
            Excel.Range Rng = App.Selection;//获取当前选中单元格
            Excel.Worksheet Sh = App.ActiveSheet;//获取当前选中工作表

            try
            {
                if (Rng.Count == 1)
                {
                    //MessageBox.Show("请选择合并单元格再执行本工具！", "友情提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                else
                {
                    //关闭提示
                    App.DisplayAlerts = false;
                    //关闭屏幕更新
                    App.ScreenUpdating = false;
                    foreach (Excel.Range Rngs in App.Intersect(Rng, Sh.UsedRange))
                    {
                        String cell;
                        dynamic value;
                        if (Rngs.MergeCells == true)
                        {
                            value = Rngs.Value;
                            cell = Rngs.MergeArea.Address;
                            Rngs.UnMerge();
                            App.Range[cell].Value = value;

                            //不可以这样写：
                            //Rng[cell].Value = value;
                            //Range[cell].Value = value;
                        }
                    }

                    //恢复提示
                    App.DisplayAlerts = true;
                    //恢复屏幕更新
                    App.ScreenUpdating = true;
                }
            }
            catch (Exception)
            {
                //MessageBox.Show("拆分单元格时遇到了错误！！！", "Excel", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
        #endregion

        #region 数据转换：合并单元格
        private void splitButton1_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Application App = Globals.ThisAddIn.Application;
            Range Rng = App.Selection;//获取当前选中单元格
            Worksheet Sh = App.ActiveSheet;//获取当前选中工作表

            #region 方法2
            
            //关闭提示
            App.DisplayAlerts = false;
            //关闭屏幕更新
            App.ScreenUpdating = false;

            ////获取选中的开始行列
            //int startRow = Rng.Row;//开始行
            //int startColumn = Rng.Column;//开始列
            //int endRow = Rng.Rows.Count + startRow;//结束行
            //int endColumn = Rng.Columns.Count + startColumn;//结束列

            ////返回选中列最大行
            //int MaxRow = ((Excel.Range)(Sh.Cells[Sh.Rows.Count, Rng.Column])).End[Excel.XlDirection.xlUp].Row;

            ////返回选中区域最大行
            //int MaxRow2 = Rng.Rows.Count + startRow - 1;

            // 获取选中单元格内容到数组
            object[,] listArr = Rng.Value;
            
            for (int i = 1; i <= Rng.Rows.Count; i++)
            {
                for (int j = 1; j <= Rng.Columns.Count; j++)
                {
                    string value = Excel_Class.类型转换_到文本(listArr[i, j]);
                    string i_value = i >= Rng.Rows.Count ? "&#$%!$DAD" : Excel_Class.类型转换_到文本(listArr[i + 1, j]);
                    string j_value = j >= Rng.Columns.Count ? "&#$&#$%$D" : Excel_Class.类型转换_到文本(listArr[i, j + 1]);
                    
                    if (Rng.Cells[i, j].MergeCells == true)
                    {
                        string cell = Rng.Cells[i, j].MergeArea.Address;
                        value = Excel_Class.类型转换_到文本(Sh.Range[cell.Substring(0, cell.IndexOf(":"))].Value);

                        if (value == i_value)
                        {
                            Sh.Range[cell, Rng.Cells[i + 1, j]].Merge();
                        }
                        else if (value == j_value)
                        {
                            Sh.Range[cell, Rng.Cells[i, j + 1]].Merge();
                        }
                    }
                    else
                    {
                        if (value == i_value)
                        {
                            Sh.Range[Rng.Cells[i, j], Rng.Cells[i + 1, j]].Merge();
                        }
                        else if (value == j_value)
                        {
                            Sh.Range[Rng.Cells[i, j], Rng.Cells[i, j + 1]].Merge();
                        }
                    }
                }
            }
            //居中单元格
            Rng.VerticalAlignment = -4108;
            Rng.HorizontalAlignment = -4108;
            //恢复提示
            App.DisplayAlerts = true;
            //恢复屏幕更新
            App.ScreenUpdating = true;


            
            #endregion
        }
        #endregion

        #region 帮助 - 开始

        #region 帮助-官网
        private void button79_Click_1(object sender, RibbonControlEventArgs e)
        {
            // 跳转到网页
            string url = "https://90le.cn/api/DTI_Tool/index.html";  // 将此URL替换为您要跳转的网页
            try
            {
                Process.Start(url);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"无法打开浏览器。\n请手动前往：https://90le.cn/api/DTI_Tool/index.html", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        #endregion

        #region 帮助-关注
        private void button128_Click_1(object sender, RibbonControlEventArgs e)
        {
            FocusOn OK_FocusOn = null;

            if (OK_FocusOn == null || OK_FocusOn.IsDisposed)
            {
                OK_FocusOn = new FocusOn();
                IntPtr handle = Process.GetCurrentProcess().MainWindowHandle;
                NativeWindow win = NativeWindow.FromHandle(handle);
                OK_FocusOn.Show();
            }
        }
        #endregion

        #region 帮助-设置
        private void button127_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Application App = Globals.ThisAddIn.Application;
            Range Rng = App.Selection;//获取当前选中单元格
            Worksheet Sh = App.ActiveSheet;//获取当前选中工作表

            //Sh.Range["A1"].Value = "CPU编号";
            //Sh.Range["B1"].Value = 系统.系统_获得CPU编号();

            //Sh.Range["A2"].Value = "硬盘序列号";
            //Sh.Range["B2"].Value = 系统.系统_获取硬盘序列号();

            //Sh.Range["A3"].Value = "硬盘的大小";
            //Sh.Range["B3"].Value = 系统.系统_获取硬盘的大小();

            //Sh.Range["A4"].Value = "网卡硬件地址";
            //Sh.Range["B4"].Value = 系统.系统_获取网卡硬件地址();

            //Sh.Range["A5"].Value = "获取IP地址";
            //Sh.Range["B5"].Value = 系统.系统_获取IP地址();

            //Sh.Range["A6"].Value = "获取计算机名";
            //Sh.Range["B6"].Value = 系统.系统_获取计算机名();

            //Sh.Range["A7"].Value = "取操作系统类型";
            //Sh.Range["B7"].Value = 系统.系统_取操作系统类型();

            //Sh.Range["A8"].Value = "显卡PNPDeviceID";
            //Sh.Range["B8"].Value = 系统.系统_显卡PNPDeviceID();

            //Sh.Range["A9"].Value = "声卡PNPDeviceID";
            //Sh.Range["B9"].Value = 系统.系统_声卡PNPDeviceID();

            //Sh.Range["A10"].Value = "CPU版本信息";
            //Sh.Range["B10"].Value = 系统.系统_CPU版本信息();

            //Sh.Range["A11"].Value = "CPU名称信息";
            //Sh.Range["B11"].Value = 系统.系统_CPU名称信息();

            //Sh.Range["A12"].Value = "CPU制造商";
            //Sh.Range["B12"].Value = 系统.系统_CPU制造商();

            //Sh.Range["A13"].Value = "主板制造商";
            //Sh.Range["B13"].Value = 系统.系统_主板制造商();

            //Sh.Range["A14"].Value = "主板编号";
            //Sh.Range["B14"].Value = 系统.系统_主板编号();

            //Sh.Range["A15"].Value = "主板型号";
            //Sh.Range["B15"].Value = 系统.系统_主板型号();

            //Sh.Range["A16"].Value = "是否64位";
            //Sh.Range["B16"].Value = 系统.系统_是否64位();

            //Sh.Range["A17"].Value = "取计算机名";
            //Sh.Range["B17"].Value = 系统.系统_取计算机名();

            //Sh.Range["A18"].Value = "取登录用户名";
            //Sh.Range["B18"].Value = 系统.系统_取登录用户名();
        }
        # endregion 帮助 - 结束

        #endregion

        #region 关键词上色
        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            DTI_Forms.Label_Color Label_Color = null;
            if (Label_Color == null || Label_Color.IsDisposed)
            {
                Label_Color = new DTI_Forms.Label_Color();
                // 获取Excel应用程序对象
                Excel.Application excelApp = Globals.ThisAddIn.Application;
                // 获取当前工作簿窗口
                Excel.Window activeWindow = excelApp.ActiveWindow;
                // 获取当前工作簿窗口的句柄
                IntPtr hwnd = new IntPtr(activeWindow.Hwnd);
                // 设置窗口置顶
                Label_Color.TopMost = true;
                // 将窗口绑定到Excel当前窗口
                Win32API.SetParent(Label_Color.Handle, hwnd);
                Label_Color.Show();
            }

        }

        #endregion

        #region 提取关键词
        private void button3_Click(object sender, RibbonControlEventArgs e)
        {
            DTI_Forms.Lable_KeyWord Lable_KeyWord = null;
            if (Lable_KeyWord == null || Lable_KeyWord.IsDisposed)
            {
                Lable_KeyWord = new DTI_Forms.Lable_KeyWord();
                IntPtr handle = Process.GetCurrentProcess().MainWindowHandle;
                NativeWindow win = NativeWindow.FromHandle(handle);
                Lable_KeyWord.Show();
            }
        }
        #endregion


        private void button5_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Application App = Globals.ThisAddIn.Application;
            Excel.Range Rng = App.Selection;//获取当前选中单元格
            Excel.Worksheet Sh = App.ActiveSheet;//获取当前选中工作表

            try
            {
                DialogResult result = MessageBox.Show("确定对当前选定区域进行加密操作吗？此操作不可撤销！", "数值转换", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);
                if (result == DialogResult.OK)
                {
                    Object[,] arr = Rng.Value2;
                    int TotalSize = arr.GetLength(0);//获取总内容行数
                    int Size = TotalSize <= 200 ? TotalSize : TotalSize / 50;//最多分成50个线程，每个线程需要执行的数量
                    int arrSize = TotalSize % Size == 0 ? TotalSize / Size : TotalSize / Size + 1;//共需要几个线程
                    string exTisp = null;
                    App.StatusBar = "加密中，请耐心等待...";

                    //使用线程
                    Task.Run(() =>
                    {
                        //创建线程list
                        List<Task> taskList = new List<Task>();
                        //初始处理行数
                        int countkey = 1;

                        //循环创建线程
                        for (int t = 0; t < arrSize; t++)
                        {
                            //把t存入s，保证线程变量的正确
                            int s = t;
                            //新建一个线程
                            taskList.Add(Task.Run(() =>
                            {
                                try
                                {
                                    //跟进这个线程要处理的起止数量，来循环内容数组
                                    for (int i = (s * Size) + 1; i <= Size * (s + 1); i++)
                                    {
                                        if (i <= TotalSize)
                                        {
                                            for (int j = 1; j <= arr.GetLength(1); j++)
                                            {
                                                if (arr[i, j] != null)
                                                {
                                                    arr[i, j] = String_Class.SHA256EncryptString(arr[i, j].ToString());
                                                }
                                            }
                                            //总处理行数+1
                                            countkey++;
                                        }
                                    }
                                }
                                catch (Exception ex)
                                {
                                    countkey = TotalSize;
                                    exTisp = ex.Message;
                                }
                            }));
                        }
                        //等待所有线程执行完成
                        Task.WaitAll(taskList.ToArray());
                        if (exTisp == null)
                        {
                            //替换结果至单元格内容
                            App.StatusBar = "正在存储结果，请耐心等待...";
                            Rng.Value2 = arr;
                            GC.Collect();
                            App.StatusBar = "已完成加密！";
                            Rng.Select();
                        }
                        else
                        {
                            MessageBox.Show("转换失败！！！\n异常信息：" + exTisp, "Excel", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        }
                        
                    });
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("转换失败！！！\n异常信息：" + ex.Message, "Excel", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }


        private void button15_Click(object sender, RibbonControlEventArgs e)
        {
            DTI_Forms.Text_MatchDelete Text_MatchDelete = null;
            if (Text_MatchDelete == null || Text_MatchDelete.IsDisposed)
            {
                Text_MatchDelete = new DTI_Forms.Text_MatchDelete();
                IntPtr handle = Process.GetCurrentProcess().MainWindowHandle;
                NativeWindow win = NativeWindow.FromHandle(handle);
                Text_MatchDelete.Show();
            }
        }
    }

}
