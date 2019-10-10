using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace HandlingFunction
{
    public class ZhuanYeFen
    {
        public string SchoolName { get; set; }
        public string MajorCode { get; set; }
        public string MajorName { get; set; }
        public string PiCi { get; set; }
        public string LowestScore { get; set; }
        public string LowestRank { get; set; }

        //public ZhuanYeFen()
        //{
        //    SchoolName = null;
        //    MajorCode = null;
        //    MajorName = null;
        //    PiCi = null;
        //    LowestScore = null;
        //    LowestRank = null;

        public void Run(string[] richTextbox1lines, string[] richTextbox2lines,string SchoolName,string PiCi)
        {
            ZhuanYeFen ZYF = new ZhuanYeFen();

            List<ZhuanYeFen> lstZYF = new List<ZhuanYeFen>();

            for (int i = 0; i < richTextbox1lines.Length; i++)
            {

                if (Regex.Match(richTextbox1lines[i], @"(^[\dA-Za-z]{2,3})").Length > 0)
                {
                    ZYF.MajorCode = Regex.Match(richTextbox1lines[i], @"^[\dA-Za-z]{2,3}").Value;

                    //zsjh.MajorName = Regex.Matches(richTextbox2lines[i], @"[\u4e00-\u9fa5（）、]{2,}")[0].Value;

                    try
                    {
                        ZYF.MajorName = Regex.Matches(richTextbox1lines[i], @"[\u4e00-\u9fa5][\d\D]{1,}")[0].Value;
                    }
                    catch (Exception)
                    {
                        System.Windows.Forms.MessageBox.Show("请检查专业代码和专业名称排列方式");
                        return;
                    }


                }
                else
                {
                    ZYF.MajorName = ZYF.MajorName + " " + richTextbox1lines[i];
                }

                if (i < richTextbox1lines.Length - 1 && Regex.Match(richTextbox1lines[i + 1], @"^[\dA-Za-z]{2}").Length > 0)
                {
                    lstZYF.Add(ZYF);
                    ZYF = new ZhuanYeFen();
                }
                if (i == richTextbox1lines.Length - 1)
                {
                    lstZYF.Add(ZYF);
                    ZYF = new ZhuanYeFen();
                }

            }

            if (lstZYF.Count != richTextbox2lines.Length)
            {
                MessageBox.Show("左右行数不等，请检查");
                return;
            }
            for (int i = 0; i < lstZYF.Count; i++)
            {
                if (richTextbox2lines[i].Contains("|"))
                {
                    ZYF.LowestScore = richTextbox2lines[i].Split('|')[0];
                    ZYF.LowestRank = richTextbox2lines[i].Split('|')[1];
                }
                else
                {
                    ZYF.LowestScore = richTextbox2lines[i].Substring(0, 3);
                    ZYF.LowestRank = richTextbox2lines[i].Substring(3, richTextbox2lines[i].Length - 3);
                }

                lstZYF[i].LowestScore = ZYF.LowestScore;
                lstZYF[i].LowestRank = ZYF.LowestRank;
            }

            for (int i = 0; i < lstZYF.Count; i++)
            {
                ExcelFunction.WriteExcel.WriteExcel4("Book1", SchoolName, lstZYF[i].MajorCode, lstZYF[i].MajorName, lstZYF[i].LowestRank, lstZYF[i].LowestScore,PiCi);
            }
        }
    }
}

    

