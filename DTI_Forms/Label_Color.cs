using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace DuteIT.DTI_Forms
{
    public partial class Label_Color : Form
    {
        public Label_Color()
        {
            this.StartPosition = FormStartPosition.CenterScreen;
            InitializeComponent();
        }

        private void Label_Color_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            Excel.Application App = Globals.ThisAddIn.Application;
            Excel.Range Rng = App.Selection;//获取当前选中单元格
            Excel.Worksheet Sh = App.ActiveSheet;//获取当前选中工作表
            //Excel.Characters characters;

            //关闭提示
            App.DisplayAlerts = false;
            //关闭屏幕更新
            App.ScreenUpdating = false;

            int txt_Color;
            switch (comboBox1.Text)
            {
                case "红色":
                    txt_Color = 3;
                    break;
                default:
                    txt_Color = 1;
                    break;
            }

            string key = richTextBox1.Text.Replace(",", "|");

            //try
            //{
                foreach (Excel.Range Rngs in Rng)
                {
                    if (Rngs.Value2 != null)
                    {
                        var txt_for = GetPathPoint(Rngs.Value2, key);
                        if (txt_for !=null)
                        {
                            foreach (var item in txt_for)
                            {
                                Rngs.Characters[item.Index + 1, item.Length].Font.ColorIndex = txt_Color;
                                //Rngs.Characters[item.Index+1, item.Length].Font.FontStyle = "加粗";
                            }
                        }
                    }
                }
            //}
            //catch (Exception)
            //{


            //}
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
                    list.Add(new { matchCol[i].Index, matchCol[i].Value.Length });
                }
            }
            
            return list;
        }
    }
}
