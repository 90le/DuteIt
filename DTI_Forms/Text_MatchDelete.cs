using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace DuteIT.DTI_Forms
{
    public partial class Text_MatchDelete : Form
    {
        public Text_MatchDelete()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Excel.Application App = Globals.ThisAddIn.Application;
            Excel.Range Rng = App.Selection;//获取当前选中单元格
            Excel.Workbook workbook = App.ActiveWorkbook;
            Excel.Worksheet Sh = App.ActiveSheet;//获取当前选中工作表

            App.ScreenUpdating = false; 

            try
            {
                Object[,] arr = Rng.Value2;

                int TotalSize = arr.GetLength(0);//获取总内容行数
                int Size = TotalSize <= 50 ? TotalSize : TotalSize / (TotalSize / 50);//最多分成100个线程，每个线程需要执行的数量
                int arrSize = TotalSize % Size == 0 ? TotalSize / Size : TotalSize / Size + 1;//共需要几个线程
                string exTisp = null;
                App.StatusBar = "处理中，请耐心等待...";
                Regex r = new Regex(richTextBox1.Text, RegexOptions.IgnoreCase); //不区分大小写
                List<int> line = new List<int>();
                bool radioButton1Checked = radioButton1.Checked; //单选框是否被选中

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
                                        if (countkey % 9 == 0)
                                        {
                                            try { App.StatusBar = "已处理第" + countkey.ToString() + "行内容..."; } catch (Exception) { }
                                        }

                                        for (int j = 1; j <= arr.GetLength(1); j++)
                                        {
                                            if (arr[i, j] != null)
                                            {
                                                bool isMatch = r.IsMatch(arr[i, j].ToString());
                                                if (isMatch)
                                                {
                                                    int h = Rng[i, j].Row;
                                                    if (radioButton1Checked)
                                                    {
                                                        
                                                        line.Add(h);
                                                    }
                                                    else
                                                    {
                                                        Rng.Rows[h-1].Interior.Color = 65535;
                                                    }
                                                    break;
                                                }
                                            }
                                        }
                                        //总处理行数+1
                                        countkey++;
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                TotalSize = countkey;
                                exTisp = ex.Message;
                            }
                        }));
                    }
                    //等待所有线程执行完成
                    Task.WaitAll(taskList.ToArray());
                    if (exTisp == null)
                    {
                        if (radioButton1Checked)
                        {
                            App.StatusBar = "正在删除行，请耐心等待...";
                            if (line.Count > 0)
                            {
                                line.Sort();    //升序
                                line.Reverse(); //反转排序
                                List<string> lineStr = new List<string>();
                                for (int i = 0; i < line.Count; i++)
                                {
                                    lineStr.Add(line[i].ToString() + ":" + line[i].ToString());
                                    if (i % 20 == 0 || i == line.Count - 1)
                                    {
                                        Sh.Range[string.Join(",", lineStr.ToArray())].Delete();
                                        lineStr.Clear();
                                    }
                                }
                            }
                        }
                        GC.Collect();
                        App.StatusBar = "执行结束！";
                    }
                    else
                    {
                        MessageBox.Show("转换失败！！！\n异常信息：" + exTisp, "Excel", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }

                });
            }
            catch (Exception ex)
            {
                MessageBox.Show("处理失败！！！\n异常信息：" + ex.Message, "Excel", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }

            App.ScreenUpdating = true;
        }
    }
}
