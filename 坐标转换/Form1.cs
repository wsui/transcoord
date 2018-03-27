using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Text.RegularExpressions;
using System.Threading;

namespace 坐标转换
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        string fpath = "";  //文件路径
        int zb = 1;         //1表示百度转wgs84,2表示百度转火星,3表示火星转WGS84,4表示火星转百度,5表示WGS84转百度,6表示WGS84转火星
        private void button1_Click(object sender, EventArgs e)
        {
            openFileDialog1.Title = "选择需要转换的csv文件";
            openFileDialog1.Multiselect = false;
            openFileDialog1.Filter = "csv|*.csv";
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                fpath = openFileDialog1.FileName;
                add_item();
            }
        }

        private void button1_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                String path = ((System.Array)e.Data.GetData(DataFormats.FileDrop)).GetValue(0).ToString();
                if (Path.GetExtension(path) == ".csv")
                {
                    e.Effect = DragDropEffects.Link;
                }
            }
            else
            {
                e.Effect = DragDropEffects.None;
            }
        }

        private void button1_DragDrop(object sender, DragEventArgs e)
        {
            fpath = ((System.Array)e.Data.GetData(DataFormats.FileDrop)).GetValue(0).ToString();
            add_item();
        }
        private void add_item()
        {
            button1.Text = fpath;
            button1.Font = new Font("微软雅黑", 12);
            button1.BackColor = Color.DarkCyan;
            StreamReader sr = new StreamReader(fpath);
            string line = sr.ReadLine();
            sr.Close();
            string[] ls = line.Split(',');
            comboBox1.SelectedItem = null;
            comboBox2.SelectedItem = null;
            comboBox1.Text = "请选择经度";
            comboBox2.Text = "请选择纬度";
            comboBox1.Items.Clear();
            comboBox2.Items.Clear();
            button2.Enabled = false;
            button2.BackColor = Color.LightGray;
            for (int i = 0; i < ls.Length; i++)
            {
                comboBox1.Items.Insert(i, ls[i]);
                comboBox2.Items.Insert(i, ls[i]);
            }
        }

        private void comboBox1_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            if (comboBox1.SelectedItem != null && comboBox2.SelectedItem != null)
            {
                button2.Enabled = true;
                button2.BackColor = Color.DarkCyan;
            }
        }

        private void comboBox2_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            if (comboBox1.SelectedItem != null && comboBox2.SelectedItem != null)
            {
                button2.Enabled = true;
                button2.BackColor = Color.DarkCyan;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string fname = Path.GetFileNameWithoutExtension(fpath);  //输出文件名
            string tt = "";     //添加的列名
            double[] lnglat = { 0, 0 };    //经纬度
            if (radioButton1.Checked == true)
            {
                zb = 1;
                fname += "_BaiduToWgs84_";
                tt = "wgs84_lng,wgs84_lat";
            }
            if (radioButton2.Checked == true)
            {
                fname += "_BaiduToGcj02_";
                zb = 2;
                tt = "gcj02_lng,gcj02_lat";
            }
            if (radioButton3.Checked == true)
            {
                fname += "_Gcj02ToWgs84_";
                zb = 3;
                tt = "wgs84_lng,wgs84_lat";
            }
            if (radioButton4.Checked == true)
            {
                fname += "_Gcj02ToBaidu_";
                zb = 4;
                tt = "baidu_lng,baidu_lat";
            }
            if (radioButton5.Checked == true)
            {
                fname += "_Wgs84ToBaidu_";
                zb = 5;
                tt = "baidu_lng,baidu_lat";
            }
            if (radioButton6.Checked == true)
            {
                fname += "_Wgs84ToGcj02_";
                zb = 6;
                tt = "gcj02_lng,gcj02_lat";
            }
            fname += DateTime.Now.ToString("yyyyMMddhhmmss") + ".csv";
            int lng_i = comboBox1.SelectedIndex;    //经度所在列
            int lat_i = comboBox2.SelectedIndex;    //纬度所在列
            if (lng_i == lat_i)
            {
                MessageBox.Show("请选择不同的列！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
          
            StreamReader sr = new StreamReader(fpath);
            string line = sr.ReadLine();
            int n = 0;  //文件行数
            progressBar1.Value = 0;
            progressBar1.Visible = true;
            while (line != null)
            {
                n += 1;
                line = sr.ReadLine();
            }
            sr.Close();
            //sr.BaseStream.Seek(0, SeekOrigin.Begin);
            sr = new StreamReader(fpath);
            StreamWriter sw = new StreamWriter(Path.GetDirectoryName(fpath) + "\\" + fname);
            line = sr.ReadLine();
            sw.Write(line + "," + tt + "\n");
            line = sr.ReadLine();
            int c = n / 100;  //百分之一的行数
            n = 0;
            while (line != null)
            {
                string[] ls = line.Split(',');
                if (!IsNumber(ls[lng_i]) || !IsNumber(ls[lat_i]))
                {
                    sr.Close();
                    sw.Close();
                    progressBar1.Visible = false;
                    if (File.Exists(Path.GetDirectoryName(fpath) + "\\" + fname))
                    {
                        File.Delete(Path.GetDirectoryName(fpath) + "\\" + fname);
                    }
                    MessageBox.Show("选择的列不是数字!", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                if (Convert.ToDouble(ls[lng_i]) > 180 || Convert.ToDouble(ls[lng_i]) < -180 || Convert.ToDouble(ls[lat_i]) > 90 || Convert.ToDouble(ls[lat_i]) < -90)
                {
                    sr.Close();
                    sw.Close();
                    progressBar1.Visible = false;
                    if (File.Exists(Path.GetDirectoryName(fpath) + "\\" + fname))
                    {
                        File.Delete(Path.GetDirectoryName(fpath) + "\\" + fname);
                    }
                    MessageBox.Show("输入的坐标超出范围!", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                //百度转WGS84
                if (zb == 1)
                {
                    lnglat = transfer.bd09towgs84(Convert.ToDouble(ls[lng_i]), Convert.ToDouble(ls[lat_i]));
                }
                //百度转火星
                if (zb == 2)
                {
                    lnglat = transfer.bd09togcj02(Convert.ToDouble(ls[lng_i]), Convert.ToDouble(ls[lat_i]));
                }
                //火星转WGS84
                if (zb == 3)
                {
                    lnglat = transfer.gcj02towgs84(Convert.ToDouble(ls[lng_i]), Convert.ToDouble(ls[lat_i]));
                }
                //火星转百度
                if (zb == 4)
                {
                    lnglat = transfer.gcj02tobd09(Convert.ToDouble(ls[lng_i]), Convert.ToDouble(ls[lat_i]));
                }
                //WGS84转百度
                if (zb == 5)
                {
                    lnglat = transfer.wgs84tobd09(Convert.ToDouble(ls[lng_i]), Convert.ToDouble(ls[lat_i]));
                }
                //WGS84转火星
                if (zb == 6)
                {
                    lnglat = transfer.wgs84togcj02(Convert.ToDouble(ls[lng_i]), Convert.ToDouble(ls[lat_i]));
                }
                sw.Write(line + "," + lnglat[0] + "," + lnglat[1] + "\n");
                line = sr.ReadLine();
                if (c == 0)
                {
                    progressBar1.Value = 100;
                    Application.DoEvents();
                }
                else if (n % c == 0 && progressBar1.Value < 100)
                {
                    progressBar1.Value++;
                    Application.DoEvents();
                }
                n++;
            }
            sw.Close();
            sr.Close();
            DialogResult dr = MessageBox.Show("转换成功！\n\n*******************************************************\n转换后的文件名为：" + fname
                + "\n转换后坐标添加在最后两列：" + tt + "\n********************************************************\n\n打开文件所在位置？",
                "提示", MessageBoxButtons.YesNo);
            if (dr == DialogResult.Yes)
            {
                System.Diagnostics.Process.Start(Path.GetDirectoryName(fpath));
            }
            progressBar1.Visible = false;
        }

        public bool IsNumber(String strNumber)
        {
            Regex objNotNumberPattern = new Regex("[^0-9.-]");
            Regex objTwoDotPattern = new Regex("[0-9]*[.][0-9]*[.][0-9]*");
            Regex objTwoMinusPattern = new Regex("[0-9]*[-][0-9]*[-][0-9]*");
            String strValidRealPattern = "^([-]|[.]|[-.]|[0-9])[0-9]*[.]*[0-9]+$";
            String strValidIntegerPattern = "^([-]|[0-9])[0-9]*$";
            Regex objNumberPattern = new Regex("(" + strValidRealPattern + ")|(" + strValidIntegerPattern + ")");

            return !objNotNumberPattern.IsMatch(strNumber) &&
                   !objTwoDotPattern.IsMatch(strNumber) &&
                   !objTwoMinusPattern.IsMatch(strNumber) &&
                   objNumberPattern.IsMatch(strNumber);
        }
    }
}
