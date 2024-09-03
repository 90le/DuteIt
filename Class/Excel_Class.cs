using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

//字符处理类
namespace DuteIT
{
    public static class Excel_Class
    {
        public static String 类型转换_到文本(object s)
        {
            return (s == null ? "" : s.ToString());
        }

        public static void 打标签() 
        {
            
        }

        /// <summary>
        /// 根据数量对数组进行分组（in查询不能超过1000个条目）
        /// </summary>
        /// <param name="list"></param>
        /// <param name="size">数量</param>
        public static List<List<string>> GroupArrayBySize(object[,] list, int size)
        {
            List<List<string>> listArr = new List<List<string>>();

            int arrSize = list.GetLength(0) % size == 0 ? list.GetLength(0) / size : list.GetLength(0) / size + 1;
            for (int i = 0; i < arrSize; i++)
            {
                List<string> sub = new List<string>();
                for (int j = i * size; j <= size * (i + 1) - 1; j++)
                {
                    if (j <= list.GetLength(0) - 1)
                    {
                        sub.Add(list[j,1].ToString());
                    }
                }
                listArr.Add(sub);
            }
            return listArr;
        }
    }
}
