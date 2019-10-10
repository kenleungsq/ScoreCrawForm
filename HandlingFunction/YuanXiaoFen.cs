using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace HandlingFunction
{
    public class YuanXiaoFen
    {
        public void Run(string[] richTextbox1lines, string SchoolName,string pici)
        {
            //if (richTextbox1lines.Length == 2)
            //{
            //    ExcelFunction.WriteExcel.WriteExcel4("Book1", SchoolName, richTextbox1lines[0], richTextbox1lines[1], "", "", pici);
            //}

            if (richTextbox1lines.Length > 2)
            {
                string temp = null;
                for (int i = 0; i < richTextbox1lines.Length; i++)
                {
                    temp += richTextbox1lines[i];
                }

                if (!temp.Contains("专科"))
                {
                    MessageBox.Show("请确认是录入是否专科批次");
                }
                ExcelFunction.WriteExcel.WriteExcel4("Book1", SchoolName, richTextbox1lines[0], richTextbox1lines[1], richTextbox1lines[richTextbox1lines.Length-2], "", pici);
            }
        }
    }
}
