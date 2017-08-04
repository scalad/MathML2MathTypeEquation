using System;
using System.Data;
using System.IO;
using Microsoft.Office.Interop;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using MTSDKDN;
using Word = Microsoft.Office.Interop.Word;
using IDataObject = System.Runtime.InteropServices.ComTypes.IDataObject;
using System.Text.RegularExpressions;
using System.Linq.Expressions;

namespace ConvertEquations
{
	/// Program Class
	#region Program Class
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
            string path = Utils.GetInputFolder("MathML8.txt");
            string d = System.IO.File.ReadAllText(path);

            //create document
            Word.Application newapp = new Word.Application();
            //用于作为函数的默认参数
            object nothing = System.Reflection.Missing.Value;
            //create a word document
            Word.Document newdoc = newapp.Documents.Add(ref nothing, ref nothing, ref nothing, ref nothing);
            //是否显示word程序界面
            newapp.Visible = false;

            string[] oneLevelData = d.Split(new string[] { "<math", "</math>" },StringSplitOptions.None);

            foreach (string datas in oneLevelData)
            {
                Word.Paragraph data = newdoc.Content.Paragraphs.Add(nothing);

                if (datas.StartsWith(" xmlns="))
                {
                    // MML in a text file to clipboard text
                    ce.Convert(new EquationInputFileText("<math" + datas + "</math>", ClipboardFormats.cfMML), new EquationOutputClipboardText());
                    try
                    {
                        Object objUnit = Word.WdUnits.wdStory;
                        newapp.Selection.EndKey(ref objUnit);
                        newdoc.Select();
                        data.Range.Paste();
                        //移动焦点并换行  
                        object count = 14;
                        newapp.Selection.MoveDown(ref nothing, ref count, ref nothing);//移动焦点
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine(e);
                    }
                    finally
                    {
                        Utils.killAllProcess("mathtype.exe");
                    }
                }
                else
                {
                    Console.WriteLine(datas);
                    
                    Regex regex = new Regex(Common.HTML_IMG, RegexOptions.IgnoreCase);
                    MatchCollection matches = regex.Matches(datas.ToString());
                    int i = 0;
                    string[] sUrlList = new string[matches.Count];
                    foreach (Match match in matches)
                    {
                        sUrlList[i++] = match.Groups["imgUrl"].Value;
                        Console.WriteLine(sUrlList[0]);
                        object LinkToFile = false;
                        object SaveWithDocument = true;
                        object Anchor = newdoc.Application.Selection.Range;

                        //插入图片
                        newdoc.Application.ActiveDocument.InlineShapes.AddPicture(sUrlList[0], ref LinkToFile, ref SaveWithDocument, ref Anchor);
                    }

                    string para = Utils.NoHTML(datas);
                    if (para != null && para != "")
                    {
                        newapp.Selection.TypeText(para);
                    }
                }
                //保存文档
                newdoc.SaveAs(ref name, ref nothing, ref nothing, ref nothing, ref nothing, ref nothing, ref nothing,
                       ref nothing, ref nothing, ref nothing, ref nothing, ref nothing, ref nothing, ref nothing,
                       ref nothing, ref nothing);
                //清空粘贴板，否则会将前一次粘贴记录保留。
                Clipboard.SetDataObject("", true);
            }
            
            //关闭文档
            object saveOption = Word.WdSaveOptions.wdDoNotSaveChanges;
            newdoc.Close(ref nothing, ref nothing, ref nothing);
            newapp.Application.Quit(ref saveOption, ref nothing, ref nothing);
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
  
            Console.ReadLine();
		}
	}
	#endregion
}
