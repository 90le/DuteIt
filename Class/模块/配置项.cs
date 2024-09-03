using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

namespace DuteIT
{
    public class 配置项
    {
        #region API写配置项

        [DllImport("kernel32")]
        //                        读配置文件方法的6个参数：所在的分区（section）、键值、     初始缺省值、     StringBuilder、   参数长度上限、配置文件路径
        private static extern int GetPrivateProfileString(string section, string key, string deVal, StringBuilder retVal,
            int size, string filePath);

        [DllImport("kernel32")]
        //                            写配置文件方法的4个参数：所在的分区（section）、  键值、     参数值、        配置文件路径
        private static extern long WritePrivateProfileString(string section, string key, string val, string filePath);

        public static void 写配置项(string section, string key, string value)
        {
            //写配置项参数

            string strPath = Environment.CurrentDirectory + "\\config.ini";
            WritePrivateProfileString(section, key, value, strPath);
        }

        public static string 读配置项(string section, string key)
        {
            StringBuilder sb = new StringBuilder(255);
            string strPath = Environment.CurrentDirectory + "\\config.ini";
            //最好初始缺省值设置为非空，因为如果配置文件不存在，取不到值，程序也不会报错
            GetPrivateProfileString(section, key, null, sb, 255, strPath);
            return sb.ToString();

        }
        #endregion

        #region 配置项,索引替换
        /// <summary>
        /// 配置项_替换
        /// </summary>
        /// <param name="strFilePath">txt等文件的路径</param>
        /// <param name="strIndex">索引的字符串，定位到某一行</param>
        /// <param name="newValue">替换新值</param>
        public static void 配置项_替换(string strFilePath, string strIndex, string newValue)
        {
            if (File.Exists(strFilePath))
            {
                string[] lines = System.IO.File.ReadAllLines(strFilePath);
                for (int i = 0; i < lines.Length; i++)
                {
                    if (lines[i].Contains(strIndex))
                    {
                        string[] str = lines[i].Split('=');
                        str[1] = newValue;
                        lines[i] = str[0] + "= " + str[1];
                    }
                }
                File.WriteAllLines(strFilePath, lines);
            }
        }
        #endregion

    }
}
