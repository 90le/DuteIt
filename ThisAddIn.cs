using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;

namespace DuteIT
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            /**
             * 获取插件安装的目录，启动时加载xlam加载宏。
             * 
             */
            //获取插件DLL的所在位置
            //string strPath = System.AppDomain.CurrentDomain.BaseDirectory;

            //根据xlam文件的路径（存放在VSO项目发布之后的根目录下）
            //string strXlam = System.IO.Path.Combine(strPath, "dtiFUN.xlam");

            //加载宏
            //Globals.ThisAddIn.Application.Workbooks.Open(strXlam);


            //不能在这里注册加载宏的函数描述，需要进去后在进行注册。
            /*
            Globals.ThisAddIn.Application.MacroOptions2(
                Macro: "test", 
                Description: "函数描述", 
                Category: "函数分组", 
                ArgumentDescriptions:new[] { "参数1","参数2"}
            );*/
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO 生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
