using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using NPOI.SS.Formula.Functions;
using Sunny.UI;
using Sunny.UI.Win32;

namespace Lean.Scanning
{
    public partial class Scanning_InScan : UIForm
    {
        public Scanning_InScan()
        {
            InitializeComponent();
        }

        private void Scanning_InScan_Load(object sender, EventArgs e)
        {
            this.TxtMsg.Visible = false;
        }
        public String DTAConnectionString = ConfigurationManager.ConnectionStrings["SerialDB"].ToString();
        public string strsn;
        public int strlen;
        private void txtINserialno_KeyPress(object sender, KeyPressEventArgs e)
        {
            //if ((e.KeyChar < (char)48) || (e.KeyChar > (char)57))
            //{
            //    e.Handled = true;
            //}
            if (txtINserialno.Text.Length > 0)
            {
                // 將 txtLot.Text 的中文字刪除
                for (int i = txtINserialno.Text.Length - 1; i >= 0; i--)
                {
                    if (!(System.Text.RegularExpressions.Regex.IsMatch(txtINserialno.Text.Substring(i, 1), @"^\w+$")))
                    {
                        txtINserialno.Text = txtINserialno.Text.Remove(i, 1);
                    }
                }
                txtINserialno.SelectionStart = txtINserialno.Text.Length;
            }


            //将用户输入的小写字母转成大写的形式
            e.KeyChar = char.ToUpper(e.KeyChar);
        }

        private void txtINserialno_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                SqlConnection DTAConnection = new SqlConnection(DTAConnectionString);
                if (e.KeyCode == Keys.Enter)
                {
                    if (this.txtINserialno.Text.Length == 0 || this.txtINserialno.Text.Length < 19)
                    {
                        if (DTAConnection.State == ConnectionState.Closed)
                        {
                            DTAConnection.Open();
                        }
                        Sound.Play("error");
                        txtINserialno.Visible = Visible;
                        TxtMsg.Text = "确认条码正确与否后，请解锁。" + this.txtINserialno.Text;
                        SqlCommand InsproaSqlcom = new SqlCommand("INSERT INTO [LeanSerial].[dbo].[DTASSET_ERROR] (errorstno,errordate,stuserid,sthostname,sthostip,sthostmac,stmemo)VALUES  ('" + this.txtINserialno.Text + "','" + DateTime.Now.ToString("yyyyMMddHHmmss") + "','AppCube','" + Get_Info.GetComputerName() + "','" + Get_Info.GetIPAddress() + "','" + Get_Info.GetMacAddress() + "','InError') ", DTAConnection);
                        InsproaSqlcom.ExecuteNonQuery();
                        txtINserialno.Clear();
                        txtINserialno.Visible = false;
                        TxtMsg.Visible = true;
                        //MessageBox.Show(this.txtproSN.Text.Substring(0, 10) + str);
                        //MessageBox.Show("序号位数不正确，请确认。");
                        //RbtnS.Focus();
                        if (Convert.ToInt32(CountS.Text) == 0)
                        {
                            CountS.Text = (Convert.ToInt32(CountS.Text)).ToString();
                        }
                        else
                        {
                            CountS.Text = (Convert.ToInt32(CountS.Text) - Convert.ToInt32(strsn)).ToString();
                        }
                        return;
                    }
                    else
                    {
                        try
                        {
                            string DR = this.txtINserialno.Text;
                            DataTable DTAmb = new DataTable();
                            string DTAmbSQL = "SELECT *  FROM [LeanSerial].[dbo].[DTASSET_INVMB] WHERE MB001='" + txtINserialno.Text.Substring(0, 10) + "' ";
                            SqlDataAdapter DTAmbAdapter = new SqlDataAdapter(DTAmbSQL, DTAConnectionString);
                            DTAmbAdapter.Fill(DTAmb);
                            if (DTAmb.Rows.Count == 0)
                            {
                                if (DTAConnection.State == ConnectionState.Closed)
                                {
                                    DTAConnection.Open();
                                }
                                Sound.Play("error");
                                txtINserialno.Visible = Visible;
                                TxtMsg.Text = "品号不存在，请解锁。" + this.txtINserialno.Text;
                                SqlCommand InsproaSqlcom = new SqlCommand("INSERT INTO [LeanSerial].[dbo].[DTASSET_ERROR] (errorstno,errordate,stuserid,sthostname,sthostip,sthostmac,stmemo)VALUES  ('" + this.txtINserialno.Text + "','" + DateTime.Now.ToString("yyyyMMddHHmmss") + "','AppCube','" + Get_Info.GetComputerName() + "','" + Get_Info.GetIPAddress() + "','" + Get_Info.GetMacAddress() + "','InError') ", DTAConnection);
                                InsproaSqlcom.ExecuteNonQuery();
                                txtINserialno.Clear();
                                txtINserialno.Visible = false;
                                TxtMsg.Visible = true;
                                if (DTAConnection.State == ConnectionState.Open)
                                {
                                    DTAConnection.Close();
                                }
                                return;
                            }
                            else
                            {


                                DataTable DTAStockTable = new DataTable();
                                string DTAStockSQL = "SELECT INSERIAL  FROM [LeanSerial].[dbo].[DTASSET_SCANNER_ALL] WHERE INSERIAL='" + txtINserialno.Text + "' ";
                                SqlDataAdapter DTAStockAdapter = new SqlDataAdapter(DTAStockSQL, DTAConnectionString);
                                DTAStockAdapter.Fill(DTAStockTable);
                                //显示品名
                                if (DTAStockTable.Rows.Count == 0)
                                {
                                    //获取品号基本信息            

                                    if (DTAConnection.State == ConnectionState.Closed)
                                    {
                                        DTAConnection.Open();
                                    }



                                    strlen = this.txtINserialno.Text.Length;
                                    if (Convert.ToInt32(this.txtINserialno.Text.Substring(strlen - 2, 2)) == 0)
                                    {
                                        strsn = this.txtINserialno.Text.Substring(strlen - 3, 3);

                                    }
                                    else
                                    {
                                        strsn = this.txtINserialno.Text.Substring(strlen - 2, 2);
                                    }
                                    SqlCommand InsproaSqlcom = new SqlCommand("INSERT INTO [LeanSerial].[dbo].[DTASSET_SCANNER_ALL] ([INSERIAL],[INDATE],[INQTY],[OUTINVOICE],[OUTDESC],[OUTSERIAL],[OUTDATE],[OUTQTY],[OGION],[INSYSUID],[INHNAME],[INHIP],[INHMAC],[INUSER],[INDTIME],[OUTSYSUID],[OUTHNAME],[OUTHIP],[OUTHMAC],[OUTUSER],[OUTDTIME],[studfchar1],[studfchar2],[studfchar3],[studfchar4],[studfint1],[studfint2],[studfint3],[studfint4] )" +
                                    " VALUES  ('" + txtINserialno.Text + "','" + DateTime.Now.ToString("yyyyMMdd") + "'," + strsn + ", '','','','','0','','InStock','" + Get_Info.GetComputerName() + "','" + Get_Info.GetIPAddress() + "','" + Get_Info.GetMacAddress() + "','USER','" + DateTime.Now.ToString("yyyyMMddHHmmss") + "','','','','','','','','','','','0','0','0','0') ", DTAConnection);
                                    InsproaSqlcom.ExecuteNonQuery();


                                    CountS.Text = (Convert.ToInt32(CountS.Text) + Convert.ToInt32(strsn)).ToString();
                                    //CountS.Text = (Convert.ToInt32(CountS.Text) + snstr).ToString();
                                    CountS.Visible = true;
                                    //Sound.Play("ok");
                                    //MessageBox.Show("数据添加成功！", "警告");
                                    // 将DataGridView绑定到BindingSource
                                    // 从数据库加载数据
                                    //this.txtproSN.Text = "";
                                    //this.LblParts.Text = "";
                                    //this.LblClass.Text = "";
                                    this.txtINserialno.Text = "";
                                    Sound.Play("ok");

                                    if (DTAConnection.State == ConnectionState.Open)
                                    {
                                        DTAConnection.Close();
                                    }

                                }
                                else
                                {
                                    if (DTAConnection.State == ConnectionState.Closed)
                                    {
                                        DTAConnection.Open();
                                    }
                                    Sound.Play("error");
                                    txtINserialno.Visible = Visible;
                                    TxtMsg.Text = "重复入库，请解锁。" + this.txtINserialno.Text;
                                    SqlCommand InsproaSqlcom = new SqlCommand("INSERT INTO [LeanSerial].[dbo].[DTASSET_ERROR] (errorstno,errordate,stuserid,sthostname,sthostip,sthostmac,stmemo)VALUES  ('" + this.txtINserialno.Text + "','" + DateTime.Now.ToString("yyyyMMddHHmmss") + "','AppCube','" + Get_Info.GetComputerName() + "','" + Get_Info.GetIPAddress() + "','" + Get_Info.GetMacAddress() + "','InError') ", DTAConnection);
                                    InsproaSqlcom.ExecuteNonQuery();
                                    txtINserialno.Clear();
                                    txtINserialno.Visible = false;
                                    TxtMsg.Visible = true;
                                    //MessageBox.Show(this.txtproSN.Text.Substring(0, 10) + str);
                                    //MessageBox.Show("序号位数不正确，请确认。");
                                    //RbtnS.Focus();
                                    if (Convert.ToInt32(CountS.Text) == 0)
                                    {
                                        CountS.Text = (Convert.ToInt32(CountS.Text)).ToString();
                                    }
                                    else
                                    {
                                        CountS.Text = (Convert.ToInt32(CountS.Text) - Convert.ToInt32(strsn)).ToString();
                                    }
                                    return;
                                    //错误播放声音


                                    // return;

                                }

                            }
                        }
                        catch (Exception a)
                        {
                            MessageBox.Show(a.ToString());
                        }
                    }
                }
            }
            catch (OverflowException ex)
            {
                MessageBox.Show(ex.Message);
            }
            catch (FormatException ex)
            {
                MessageBox.Show(ex.Message);
            }
            catch (ArgumentNullException ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void TxtMsg_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                if (this.TxtMsg.Text == "InstockNG")
                {
                    this.txtINserialno.Visible = true;
                    this.TxtMsg.Visible = false;
                }
            }
        }

        private void BtnRef_Click(object sender, EventArgs e)
        {
            new System.Threading.Thread((System.Threading.ThreadStart)delegate
            {
                Application.Run(new Scanning_Main());
            }).Start();
            this.Close();
        }


    }
}
