using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Windows.Forms;
using System.Text.RegularExpressions;

namespace ConvertEquations
{
    /// <summary>
    /// 全局静态变量
    /// </summary>
    class Common
    {
        //过滤HTML中img标签中的src
        public static string HTML_IMG = @"<img\b[^<>]*?\bsrc[\s\t\r\n]*=[\s\t\r\n]*[""']?[\s\t\r\n]*(?<imgUrl>[^\s\t\r\n""'<>]*)[^<>]*?/?[\s\t\r\n]*>";
    }

    /// <summary>
    /// 系统状态码
    /// </summary>
    sealed class ResultCode
    {
        public static string SUCCESS = "转换成功";
        public static string WORD_ERROR = "Word错误";
        public static string WORD_SAVE_ERROR = "Word存储错误";
        public static string EXCEL_ERROR = "Excel错误";
        public static string EXCEL_READ_ERROR = "Excel读取错误";
        public static string FILE_PATH_ERROR = "文件位置错误";
    }


    /// <summary>
    /// 系统工具
    /// </summary>
    class Utils
    {
        // 杀掉processName进程
        public static void killAllProcess(string processName)
        {
            System.Diagnostics.Process[] myPs;
            myPs = System.Diagnostics.Process.GetProcesses();
            foreach (System.Diagnostics.Process p in myPs)
            {
                if (p.Id != 0)
                {
                    try
                    {
                        if (p.Modules != null)
                            if (p.Modules.Count > 0)
                            {
                                System.Diagnostics.ProcessModule pm = p.Modules[0];
                                if (pm.ModuleName.ToLower() == processName)
                                    p.Kill();
                            }
                    }
                    catch
                    { }
                }
            }
        }

        //过滤掉HTML中标签
        public static string NoHTML(string Htmlstring)
        {
            //删除脚本
            Htmlstring = Regex.Replace(Htmlstring, @"<script[^>]*?>.*?</script>", "",RegexOptions.IgnoreCase);
            Htmlstring = Regex.Replace(Htmlstring, @"<(.[^>]*)>", "", RegexOptions.IgnoreCase);
            Htmlstring = Regex.Replace(Htmlstring, @"([\r\n])[\s]+", "", RegexOptions.IgnoreCase);
            Htmlstring = Regex.Replace(Htmlstring, @"–>", "", RegexOptions.IgnoreCase);
            Htmlstring = Regex.Replace(Htmlstring, @"<!–.*", "", RegexOptions.IgnoreCase);
            Htmlstring = Regex.Replace(Htmlstring, @"&(quot|#34);", "\"",RegexOptions.IgnoreCase);
            Htmlstring = Regex.Replace(Htmlstring, @"&(amp|#38);", "&",RegexOptions.IgnoreCase);
            Htmlstring = Regex.Replace(Htmlstring, @"&(lt|#60);", "<",RegexOptions.IgnoreCase);
            Htmlstring = Regex.Replace(Htmlstring, @"&(gt|#62);", ">",RegexOptions.IgnoreCase);
            Htmlstring = Regex.Replace(Htmlstring, @"&(nbsp|#160);", "   ",RegexOptions.IgnoreCase);
            Htmlstring = Regex.Replace(Htmlstring, @"&(iexcl|#161);", "\xa1", RegexOptions.IgnoreCase);
            Htmlstring = Regex.Replace(Htmlstring, @"&(cent|#162);", "\xa2", RegexOptions.IgnoreCase);
            Htmlstring = Regex.Replace(Htmlstring, @"&(pound|#163);", "\xa3", RegexOptions.IgnoreCase);
            Htmlstring = Regex.Replace(Htmlstring, @"&(copy|#169);", "\xa9", RegexOptions.IgnoreCase);
            Htmlstring = Regex.Replace(Htmlstring, @"&#(\d+);", "", RegexOptions.IgnoreCase);
            Htmlstring.Replace("<", "");
            Htmlstring.Replace(">", "");
            Htmlstring.Replace("\r\n", "");
            return Htmlstring;
        }

        public static string GetInputFolder(string strFile)
        {
            FileInfo fi = new FileInfo(Application.ExecutablePath);
            string strRet = System.IO.Path.Combine(fi.Directory.Parent.Parent.FullName, "Data");
            return System.IO.Path.Combine(strRet, strFile);
        }
    }

}
