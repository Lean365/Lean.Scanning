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
    public partial class Scanning_Login : UILoginForm
    {
        public Scanning_Login()
        {
            InitializeComponent();
        }

        private void Scanning_Login_Load(object sender, EventArgs e)
        {
            this.lblTitle.Text = "序列号扫描系统";
            this.lblSubText.Text = DateTime.Now.ToString("yyyy") + " © Lean365 Inc.";
            //edtUser.Text = "admin";
        }

        private void Scanning_Login_ButtonLoginClick(object sender, EventArgs e)
        {


            //打开新窗口关闭旧窗口
            new System.Threading.Thread((System.Threading.ThreadStart)delegate
            {
                Application.Run(new Scanning_Main());
            }).Start();
            this.Close();
            UIMessageTip.ShowOk("登录成功");
        }


        private void Scanning_Login_ButtonCancelClick(object sender, EventArgs e)
        {
            this.Close();
        }


    }
}
