using System;
using System.Data;
using System.IO;
using Microsoft.Office.Interop;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using MTSDKDN;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;
using IDataObject = System.Runtime.InteropServices.ComTypes.IDataObject;
using System.Text.RegularExpressions;
using System.Linq.Expressions;
using System.Collections.Generic;

namespace ConvertEquations
{
    class Program
    {
        //微软提供的可调用的API入口
        [DllImport("shell32.dll ")]
        public static extern int ShellExecute(IntPtr hwnd, String lpszOp, String lpszFile, String lpszParams, String lpszDir, int FsShowCmd);

        public void LocalToClipboard(Program p, ConvertEquation ce)
        {
            Utils.killAllProcess("winword.exe");
            Utils.killAllProcess("mathtype.exe");
            object name = "e:\\yb3.doc";

            //create document
            Word.Application newapp = new Word.Application();
            //用于作为函数的默认参数
            object nothing = System.Reflection.Missing.Value;
            //create a word document
            Word.Document newdoc = newapp.Documents.Add(ref nothing, ref nothing, ref nothing, ref nothing);
            //是否显示word程序界面
            newapp.Visible = false;

            Excel.Workbook workbook = null;
            Excel.Worksheet worksheet = null;
            string path = Utils.GetInputFolder("高中数学(必修2)(人教版A)01一章 空间几何体 (3).xls");
            object missing = System.Reflection.Missing.Value;
            Excel.Application excel = new Excel.Application();//lauch excel application  
            if (excel == null)
            {
                Console.WriteLine("Can't access excel");
                return;
            }
            excel.Visible = false; 
            excel.UserControl = true;
            // 以只读的形式打开EXCEL文件  
            workbook = excel.Application.Workbooks.Open(path, missing, true, missing, missing, missing,
             missing, missing, missing, true, missing, missing, missing, missing, missing);
            //取得第一个工作薄
            worksheet = (Excel.Worksheet)workbook.Worksheets.get_Item(1);
            //取得总记录行数   (包括标题列)
            int iRowCount = worksheet.UsedRange.Rows.Count;
            int iColCount = worksheet.UsedRange.Columns.Count;

            //生成列头
            List<string> titles = new List<String>();
            for (int i = 0; i < iColCount; i++)
            {
                var txt = ((Excel.Range)worksheet.Cells[1, i + 1]).Text.ToString();
                titles.Add(txt.ToString()+ ": ");

            }
            //生成行数据
            Excel.Range range;
            //从第二行开始
            int rowIdx = 2;
            int count = 0;
            for (int iRow = rowIdx; iRow <= iRowCount; iRow++)
            {
                for (int iCol = 1; iCol <= iColCount; iCol++)
                {
                    //插入列头
                    newapp.Selection.Font.Color = Word.WdColor.wdColorBlue;
                    newapp.Selection.TypeText(titles[iCol - 1]);
                    //得到单元格内容
                    range = (Excel.Range)worksheet.Cells[iRow, iCol];
                    string d = range.Text.ToString();
                    string[] oneLevelData = d.Split(new string[] { "<math", "</math>" }, StringSplitOptions.None);

                    List<string> imgs = new List<string>();
                    foreach (string datas in oneLevelData)
                    {
                        if (datas.StartsWith(" xmlns="))
                        {
                            // MML in a text file to clipboard text
                            ce.Convert(new EquationInputFileText("<math" + datas + "</math>", ClipboardFormats.cfMML), new EquationOutputClipboardText());
                            count++;
                            try
                            {
                                newapp.Selection.Paste();
                            }
                            catch (Exception e)
                            {
                                Console.WriteLine(e);
                            }
                            finally
                            {
                                if (count == 9)
                                {
                                    Utils.killAllProcess("mathtype.exe");
                                    count = 0;
                                }
                            }
                        }
                        else
                        {
                            Console.WriteLine(datas);
                            Regex regex = new Regex(Common.HTML_IMG, RegexOptions.IgnoreCase);
                            MatchCollection matches = regex.Matches(datas.ToString());
                            foreach (Match match in matches)
                            {
                                imgs.Add(match.Groups["imgUrl"].Value);
                            }

                            //去除HTML标签
                            string para = Utils.NoHTML(datas);
                            if (para != null && para != "")
                            {
                                newapp.Selection.Font.Color = Word.WdColor.wdColorBlack;
                                //去除空格、插入文本
                                newapp.Selection.TypeText(para.Trim());
                            }
                        }
                    }
                    foreach (string s in imgs)
                    {
                        object SaveWithDocument = true;
                        object Anchor = newdoc.Application.Selection.Range;
                        newapp.Selection.TypeParagraph();
                        //插入图片
                        newdoc.Application.ActiveDocument.InlineShapes.AddPicture(s, ref nothing, ref SaveWithDocument, ref Anchor);
                    }
                    newapp.Selection.TypeParagraph();
                }
                newapp.Selection.TypeParagraph();
                //清空粘贴板，否则会将前一次粘贴记录保留。
                Clipboard.SetDataObject("", true);
            }

            //保存文档
            newdoc.SaveAs(ref name, ref nothing, ref nothing, ref nothing, ref nothing, ref nothing, ref nothing,
                   ref nothing, ref nothing, ref nothing, ref nothing, ref nothing, ref nothing, ref nothing,
                   ref nothing, ref nothing);
            //关闭文档
            object saveOption = Word.WdSaveOptions.wdDoNotSaveChanges;
            newdoc.Close(ref nothing, ref nothing, ref nothing);
            newapp.Application.Quit(ref saveOption, ref nothing, ref nothing);
            excel.Quit();
            excel = null;
            newdoc = null;
            newapp = null;
            ShellExecute(IntPtr.Zero, "open", name.ToString(), "", "", 3);
            Console.WriteLine("Inspect the clipboard, then press any key");
        }


        [STAThread]
        static void Main(string[] args)
        {
            Program p = new Program();
            ConvertEquation ce = new ConvertEquation();
            p.LocalToClipboard(p, ce);
        }
    }
}
