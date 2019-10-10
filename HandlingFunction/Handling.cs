using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Runtime.InteropServices;

namespace HandlingFunction
{
    public class Handling
    {
        [DllImport("user32.dll")]

        public static extern int ShowWindow(

  int hwnd,

  int nCmdShow

);

        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        List<ZhaoShengJiHua> lstzsjh = new List<ZhaoShengJiHua>();
        public bool Run(string Schoolname, string[] Lines, string[] richtextbox2Lines, string picikelei)
        {
            ZhaoShengJiHua zsjh = new ZhaoShengJiHua();
            zsjh.SchoolName = Schoolname;

            string[] richTextbox2lines = Lines;
            string[] lines2 = richtextbox2Lines;

            for (int i = 0; i < richTextbox2lines.Length; i++)
            {

                if (Regex.Match(richTextbox2lines[i], @"(^[\dA-Za-z]{2,3})").Length > 0)
                {
                    zsjh.MajorCode = Regex.Match(richTextbox2lines[i], @"^[\dA-Za-z]{2,3}").Value;

                    //zsjh.MajorName = Regex.Matches(richTextbox2lines[i], @"[\u4e00-\u9fa5（）、]{2,}")[0].Value;

                    try
                    {
                        zsjh.MajorName = Regex.Matches(richTextbox2lines[i], @"[\u4e00-\u9fa5][\d\D]{1,}")[0].Value;
                    }
                    catch (Exception)
                    {
                        System.Windows.Forms.MessageBox.Show("请检查专业代码和专业名称排列方式");
                        return false;
                    }


                }
                else
                {
                    zsjh.MajorName = zsjh.MajorName + " " + richTextbox2lines[i];
                }

                if (i < richTextbox2lines.Length - 1 && Regex.Match(richTextbox2lines[i + 1], @"^[\dA-Za-z]{2}").Length > 0)
                {
                    lstzsjh.Add(zsjh);
                    zsjh = new ZhaoShengJiHua();
                }
                if (i == richTextbox2lines.Length - 1)
                {
                    lstzsjh.Add(zsjh);
                    zsjh = new ZhaoShengJiHua();
                }

            }

            if (lstzsjh.Count != richtextbox2Lines.Length)
            {
                return false;
            }
            for (int i = 0; i < lstzsjh.Count; i++)
            {
                richtextbox2Lines[i] = richtextbox2Lines[i].Replace("|", "");

                if (richtextbox2Lines[i].Length > 4)
                {
                    if (richtextbox2Lines[i].Length < 6)
                    {
                        MessageBox.Show("请检查招生计划和费用长度是否少于6个字符");
                        return false;
                    }

                    if (richtextbox2Lines[i].Length == 6)
                    {
                        zsjh.Fee = richtextbox2Lines[i].Substring(richtextbox2Lines[i].Length - 4, 4);
                    }
                    else if (richtextbox2Lines[i].Length > 6 && richtextbox2Lines[i].Contains("4"))
                    {
                        zsjh.Fee = richtextbox2Lines[i].Split('4')[1];
                    }
                    else if (richtextbox2Lines[i].Length > 6 && !richtextbox2Lines[i].Contains("4") && Regex.IsMatch(richtextbox2Lines[i], @"\d3\d"))
                    {
                        zsjh.Fee = richtextbox2Lines[i].Split('3')[1];
                    }





                    if (richtextbox2Lines[i].Length > 6 && richtextbox2Lines[i].Contains("4"))
                    {

                        zsjh.Plan = richtextbox2Lines[i].Split('4')[0];

                    }
                    else if (richtextbox2Lines[i].Length > 6 && !richtextbox2Lines[i].Contains("4") && Regex.IsMatch(richtextbox2Lines[i], @"\d3\d"))
                    {
                        zsjh.Plan = richtextbox2Lines[i].Split('3')[0];
                    }
                    else
                    {
                        zsjh.Plan = richtextbox2Lines[i].Substring(0, 1);
                    }
                }
                else if (richtextbox2Lines[i].Contains("待定"))
                {
                    zsjh.Fee = "待定";

                    if (richtextbox2Lines[i].Length == 7)
                    {
                        zsjh.Plan = richtextbox2Lines[i].Substring(0, 2);
                    }
                    else
                    {
                        zsjh.Plan = richtextbox2Lines[i].Substring(0, 1);
                    }
                }

                lstzsjh[i].Fee = zsjh.Fee;
                lstzsjh[i].Plan = zsjh.Plan;
            }

            for (int i = 0; i < lstzsjh.Count; i++)
            {
                ExcelFunction.WriteExcel.WriteExcel4("Book1", Schoolname, lstzsjh[i].MajorCode, lstzsjh[i].MajorName, lstzsjh[i].Fee, lstzsjh[i].Plan, picikelei);
            }


            return true;




        }


    }


}

public class ZhaoShengJiHua
{
    public string SchoolCode { get; set; }
    public string SchoolName { get; set; }
    public string MajorCode { get; set; }
    public string MajorName { get; set; }
    public string Fee { get; set; }
    public string Plan { get; set; }

}
