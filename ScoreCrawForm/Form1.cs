using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Text.RegularExpressions;

namespace ScoreCrawForm
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            IE.IE ieclass = new IE.IE();
            ieclass.Run();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            //IEfunction.IE ieclass = new IEfunction.IE();
            ////ieclass.GetLink();
            //ieclass.NavigateYZYLink();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            //IEfunction.IE ieclass = new IEfunction.IE();
            //ieclass.OpenLinkInExcel();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            //IEfunction.IE ieclass = new IEfunction.IE();
            //ieclass.RunWMZYZhaoShengJiHua();
        }

        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {
            label1.Text = Regex.Matches(richTextBox1.Text, @"([\dA-Za-z]{2,3})").Count.ToString();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            //if (!checkBox1.Checked && !checkBox2.Checked)
            //{
            //    MessageBox.Show("请选择科类");
            //    return;
            //}
            HandlingFunction.Handling a = new HandlingFunction.Handling();
            bool result;
            string pici = null;

            foreach (Control ctrl in this.Controls)
            {
                if (ctrl is RadioButton)
                {
                    RadioButton rb = (RadioButton)ctrl;
                    if (rb.Checked)
                    {
                        pici = rb.Text;
                        break;
                    }
                }


            }

            if (pici == null)
            {
                MessageBox.Show("请选择批次");
                return;
            }
            result = a.Run(richTextBox3.Text, richTextBox1.Lines, richTextBox2.Lines, pici);


            if (!result)
            {
                MessageBox.Show("左右行数不等，请检查");
            }



        }

        private void button6_Click(object sender, EventArgs e)
        {
            richTextBox1.Clear();
            richTextBox2.Clear();
            richTextBox3.Clear();

            foreach (Control ctrl in this.Controls)
            {
                if (ctrl is RadioButton)
                {
                    RadioButton rb = (RadioButton)ctrl;
                    if (rb.Checked)
                    {
                        rb.Checked = false;
                    }
                }
            }

        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {


        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void button7_Click(object sender, EventArgs e)
        {
            HandlingFunction.ZhuanYeFen zyf = new HandlingFunction.ZhuanYeFen();

            string pici = null;

            foreach (Control ctrl in this.Controls)
            {
                if (ctrl is RadioButton)
                {
                    RadioButton rb = (RadioButton)ctrl;
                    if (rb.Checked)
                    {
                        pici = rb.Text;
                        break;
                    }
                }
            }

            if (pici == null)
            {
                MessageBox.Show("请选择批次");
                return;
            }

            
            zyf.Run(richTextBox1.Lines,richTextBox2.Lines, richTextBox3.Text, pici);
        }

        private void richTextBox2_TextChanged(object sender, EventArgs e)
        {
            label2.Text = richTextBox2.Lines.Length.ToString();
        }

        private void button8_Click(object sender, EventArgs e)
        {
            //string pici = null;

            //foreach (Control ctrl in this.Controls)
            //{
            //    if (ctrl is RadioButton)
            //    {
            //        RadioButton rb = (RadioButton)ctrl;
            //        if (rb.Checked)
            //        {
            //            pici = rb.Text;
            //            break;
            //        }
            //    }
            //}

            //if (pici == null)
            //{
            //    MessageBox.Show("请选择批次");
            //    return;
            //}

            HandlingFunction.YuanXiaoFen yuanxiaofen = new HandlingFunction.YuanXiaoFen();
            yuanxiaofen.Run(richTextBox1.Lines, richTextBox3.Text, "");
        }
    }

    public class ZSJH
    {
        public string MajorCode { get; set; }
        public string majorname { get; set; }
    }
}
