using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Sunny.UI;
namespace Lean.Scanning
{
    public partial class Scanning_Main : UIForm
    {
        public Scanning_Main()
        {
            InitializeComponent();
        }

        private void Scanning_Main_Load(object sender, EventArgs e)
        {
            this.Text = "序列号扫描系统 " + DateTime.Now.ToString("yyyy");
            this.uiImageButton1.Text = "入库扫描";
            this.uiImageButton2.Text = "出库扫描";
            this.uiImageButton3.Text = "报表管理";
            this.uiImageButton4.Text = "基本信息";
            this.uiImageButton5.Text = "退出系统";
            uiImageButton1.ForeColor = Color.White;
            uiImageButton2.ForeColor = Color.White;
            uiImageButton3.ForeColor = Color.White;
            uiImageButton4.ForeColor = Color.White;
            uiImageButton5.ForeColor = Color.White;
        }

        private void uiImageButton1_MouseLeave(object sender, EventArgs e)
        {
            uiImageButton1.Size = new Size(128, 138);
            uiImageButton1.BackColor = Color.Transparent;
            uiImageButton1.ForeColor = Color.LightBlue;

        }

        private void uiImageButton1_MouseMove(object sender, MouseEventArgs e)
        {
            uiImageButton1.Size = new Size(128, 138);
            uiImageButton1.BackColor = Color.LightBlue;
            uiImageButton1.ForeColor = Color.White;
        }

        private void uiImageButton2_MouseLeave(object sender, EventArgs e)
        {
            uiImageButton2.Size = new Size(128, 138);
            uiImageButton2.BackColor = Color.Transparent;
            uiImageButton2.ForeColor = Color.LightBlue;
        }

        private void uiImageButton2_MouseMove(object sender, MouseEventArgs e)
        {
            uiImageButton2.Size = new Size(128, 138);
            uiImageButton2.BackColor = Color.LightBlue;
            uiImageButton2.ForeColor = Color.White;
        }

        private void uiImageButton3_MouseLeave(object sender, EventArgs e)
        {
            uiImageButton3.Size = new Size(128, 138);
            uiImageButton3.BackColor = Color.Transparent;
            uiImageButton3.ForeColor = Color.LightBlue;
        }

        private void uiImageButton3_MouseMove(object sender, MouseEventArgs e)
        {
            uiImageButton3.Size = new Size(128, 138);
            uiImageButton3.BackColor = Color.LightBlue;
            uiImageButton3.ForeColor = Color.White;
        }

        private void uiImageButton4_MouseLeave(object sender, EventArgs e)
        {
            uiImageButton4.Size = new Size(128, 138);
            uiImageButton4.BackColor = Color.Transparent;
            uiImageButton4.ForeColor = Color.LightBlue;
        }

        private void uiImageButton4_MouseMove(object sender, MouseEventArgs e)
        {
            uiImageButton4.Size = new Size(128, 138);
            uiImageButton4.BackColor = Color.LightBlue;
            uiImageButton4.ForeColor = Color.White;
        }
        private void uiImageButton5_MouseLeave(object sender, EventArgs e)
        {
            uiImageButton5.Size = new Size(128, 138);
            uiImageButton5.BackColor = Color.Transparent;
            uiImageButton5.ForeColor = Color.LightBlue;
        }
        private void uiImageButton5_MouseMove(object sender, MouseEventArgs e)
        {
            uiImageButton5.Size = new Size(128, 138);
            uiImageButton5.BackColor = Color.LightBlue;
            uiImageButton5.ForeColor = Color.White;

        }
        private void uiImageButton1_Click(object sender, EventArgs e)
        {
            //打开新窗口关闭旧窗口
            new System.Threading.Thread((System.Threading.ThreadStart)delegate
            {
                Application.Run(new Scanning_InScan());
            }).Start();
            this.Close();
        }
        private void uiImageButton2_Click(object sender, EventArgs e)
        { 
            //打开新窗口关闭旧窗口
            new System.Threading.Thread((System.Threading.ThreadStart)delegate
            {
                Application.Run(new Scanning_OutScan());
            }).Start();
            this.Close();

        }

        private void uiImageButton3_Click(object sender, EventArgs e)
        {   
            //打开新窗口关闭旧窗口
            new System.Threading.Thread((System.Threading.ThreadStart)delegate
            {
                Application.Run(new Scanning_Report());
            }).Start();
            this.Close();

        }

        private void uiImageButton4_Click(object sender, EventArgs e)
        {
            //打开新窗口关闭旧窗口
            new System.Threading.Thread((System.Threading.ThreadStart)delegate
            {
                Application.Run(new Scanning_Base());
            }).Start();
            this.Close();
        }
        private void uiImageButton5_Click(object sender, EventArgs e)
        {
            this.Close();
        }


    }
}
