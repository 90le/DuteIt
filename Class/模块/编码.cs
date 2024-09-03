using System;
using System.Text;

namespace DuteIT
{
    public static class   编码
    {
        #region Url

        /// <summary>
        /// 编码_字符串到UTF8或gbk
        /// </summary>
        /// <param name="str"></param>
        /// <param name="e">utf8 或者gbk </param>
        /// <returns></returns>
        public static string 编码_文本到utf8或gbk(this string str, Encoding e)
        {
            return System.Web.HttpUtility.UrlEncode(str, e);
        }

        /// <summary>
        /// 编码_字符串到utf8
        /// </summary>
        /// <param name="str"></param>
        /// <returns></returns>
        public static string 编码_字符串到utf8(this string str)
        {
            return System.Web.HttpUtility.UrlEncode(str, System.Text.Encoding.UTF8);
        }
        /// <summary>
        /// Url 编码_解码utf8
        /// </summary>
        /// <param name="str"></param>
        /// <returns></returns>
        public static string 编码_解码到utf8(this string str)
        {
            return System.Web.HttpUtility.UrlDecode(str, System.Text.Encoding.UTF8);
        }

        /// <summary>
        /// url 编码_解码到utf8或gbk
        /// </summary>
        /// <param name="str"></param>
        /// <param name="e"> Encoding.Utf8 或者gbk</param>
        /// <returns></returns>
        public static string 编码_解码到utf8或gbk(this string str, Encoding e)
        {
            return System.Web.HttpUtility.UrlDecode(str, e);
        }
        #endregion URL

        #region 字节集到文本

        /// <summary>
        /// 编码_UTF8转文本
        /// </summary>
        /// <param name="bin"></param>
        /// <returns></returns>
        public static string 编码_UTF8转文本(this byte[] bin)
        {
            return Encoding.UTF8.GetString(bin);
        }


        /// <summary>
        ///  编码_转文本
        /// </summary>
        /// <param name="bin"></param>
        /// <param name="e"> 编码</param>
        /// <returns></returns>
        public static string 编码_转文本(this byte[] bin, Encoding e)
        {
            return e.GetString(bin);
        }

        /// <summary>
        /// 编码_文本到节集
        /// </summary>
        /// <param name="str"></param>
        /// <returns></returns>
        public static byte[] 编码_文本到字节集(this string str)
        {

            return Encoding.Default.GetBytes(str);

        }

        /// <summary>
        /// 编码_文本到Encoding
        /// </summary>
        /// <param name="str"></param>
        /// <param name="e">Encoding 编码</param>
        /// <returns></returns>
        public static byte[] 编码_文本到Encoding(this string str, Encoding e)
        {
            return e.GetBytes(str);
        }


        #endregion 字节集到文本

        #region UniCode码字符串
        /// <summary>  
        /// 字符串转为UniCode码字符串  编码_文本到UniCode
        /// </summary>  
        /// <param name="s"></param>  
        /// <returns></returns>  
        public static string 编码_文本到UniCode(this string s)
        {
            char[] charbuffers = s.ToCharArray();
            byte[] buffer;
            StringBuilder sb = new StringBuilder();
            for (int i = 0; i < charbuffers.Length; i++)
            {
                buffer = System.Text.Encoding.Unicode.GetBytes(charbuffers[i].ToString());
                sb.Append(String.Format("\\u{0:X2}{1:X2}", buffer[1], buffer[0]));
            }
            return sb.ToString();
        }
        /// <summary>  
        /// Unicode字符串转为正常字符串   编码_UniCode到文本
        /// </summary>  
        /// <param name="srcText"></param>  
        /// <returns></returns>  
        public static string 编码_UniCode到文本(this string srcText)
        {
            string dst = "";
            string src = srcText;
            int len = srcText.Length / 6;
            for (int i = 0; i <= len - 1; i++)
            {
                string str = "";
                str = src.Substring(0, 6).Substring(2);
                src = src.Substring(6);
                byte[] bytes = new byte[2];
                bytes[1] = byte.Parse(int.Parse(str.Substring(0, 2), System.Globalization.NumberStyles.HexNumber).ToString());
                bytes[0] = byte.Parse(int.Parse(str.Substring(2, 2), System.Globalization.NumberStyles.HexNumber).ToString());
                dst += Encoding.Unicode.GetString(bytes);
            }
            return dst;
        }


        #endregion UniCode码字符串
    }
}
