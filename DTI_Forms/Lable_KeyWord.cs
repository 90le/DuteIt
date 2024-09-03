using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace DuteIT.DTI_Forms
{
    public partial class Lable_KeyWord : Form
    {
        public Lable_KeyWord()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Excel.Application App = Globals.ThisAddIn.Application;
            Excel.Range Rng = App.Selection;//获取当前选中单元格
            Excel.Worksheet Sh = App.ActiveSheet;//获取当前选中工作表

            //关闭提示
            App.DisplayAlerts = false;
            //关闭屏幕更新
            App.ScreenUpdating = false;


            string txt = richTextBox1.Text.Replace(",", "|");

            object[,] arr = Rng.Value;
            try
            {
                for (int i = 1; i <= arr.GetLength(0); i++)
                {
                    for (int j = 1; j <= arr.GetLength(1); j++)
                    {
                        var txt_for = GetPathPoint(arr[i, j].ToString(), txt);
                        if (txt_for != null)
                        {
                            arr[i, j] = String.Join(",", txt_for.ToArray());
                        }
                        else
                        {
                            arr[i, j] = null;
                        }
                    }
                }

                //Rng.Value = arr;
                Sh.Range[textBox1.Text].Resize[arr.GetLength(0), arr.GetLength(1)].Value2 = arr;
            }
            catch (Exception)
            {


            }

            //恢复提示
            App.DisplayAlerts = true;
            //恢复屏幕更新
            App.ScreenUpdating = true;
        }

        /// <summary>
        /// 获取正则表达式匹配结果集
        /// </summary>
        /// <param name="value">字符串</param>
        /// <param name="regx">正则表达式</param>
        private List<object> GetPathPoint(string value, string regx)
        {
            if (string.IsNullOrWhiteSpace(value))
                return null;
            bool isMatch = System.Text.RegularExpressions.Regex.IsMatch(value, regx);
            if (!isMatch)
                return null;
            System.Text.RegularExpressions.MatchCollection matchCol = System.Text.RegularExpressions.Regex.Matches(value, regx);
            List<object> list = new List<object>();

            if (matchCol.Count > 0)
            {
                for (int i = 0; i < matchCol.Count; i++)
                {
                    list.Add(matchCol[i].Value);
                }
            }

            return list;
        }
    }
}
