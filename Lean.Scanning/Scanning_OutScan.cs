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
    public partial class Scanning_OutScan : UIForm
    {
        public Scanning_OutScan()
        {
            InitializeComponent();
        }
        public String DTAConnectionString = ConfigurationManager.ConnectionStrings["SerialDB"].ToString();
        public string strsn;
        public int strlen;

        private void Scanning_OutScan_Load(object sender, EventArgs e)
        {
            this.TxtMsg.Visible = false;
            DataBindCmbXB();
            DataBindCmbLB();
        }
        protected void DataBindCmbLB() //为datagridview綁定数据源
        {
            /// <summary> 
            /// //为CmbDays绑定数据源
            /// </summary>


            SqlConnection TypeConnection = new SqlConnection(DTAConnectionString);
            TypeConnection.Open();
            DataTable TypeTable = new DataTable();
            string TypeSQL = "SELECT DISTINCT REPLACE(LB001,' ','') LB001,REPLACE(LB002,' ','')LB002 FROM DTASSET_INVLB";
            SqlDataAdapter TypeAdapter = new SqlDataAdapter(TypeSQL, TypeConnection);
            DataSet Typeds = new DataSet();
            TypeAdapter.Fill(Typeds);
            //在cmbDays中添加“选择提示”

            TypeTable = Typeds.Tables[0];
            DataRow Typedr = TypeTable.NewRow();
            Typedr[0] = "Desc";
            Typedr[1] = "目的地";
            TypeTable.Rows.InsertAt(Typedr, 0);

            this.cmbDesc.DataSource = TypeTable;
            cmbDesc.DisplayMember = "LB002";//显示内容
            cmbDesc.ValueMember = "LB001";

            //CmbDays.ValueMember = "ID";//选项对应的value
            //this.DgvPro1b.DataSource = Myds.Tables[0]; //把第一个表设置为数据源
            //与ASP.NET里GridView控件不同的是这里的DataGridView不需要DataBind()

            if (TypeConnection.State == ConnectionState.Open)
            {
                TypeConnection.Close();
            }
            /// <summary> 
            /// //为CmbDays绑定数据源
            /// </summary>

        }
        protected void DataBindCmbXB() //为datagridview綁定数据源
        {
            /// <summary> 
            /// //为CmbDays绑定数据源
            /// </summary>


            SqlConnection TypeConnection = new SqlConnection(DTAConnectionString);
            TypeConnection.Open();
            DataTable TypeTable = new DataTable();
            string TypeSQL = "SELECT DISTINCT REPLACE(XB001+'-'+XB002,' ','') AS XB001 FROM [LeanSerial].[dbo].[DTASSET_INVXB]";
            SqlDataAdapter TypeAdapter = new SqlDataAdapter(TypeSQL, TypeConnection);
            DataSet Typeds = new DataSet();
            TypeAdapter.Fill(Typeds);
            //在cmbDays中添加“选择提示”

            TypeTable = Typeds.Tables[0];
            DataRow Typedr = TypeTable.NewRow();
            Typedr[0] = "仕向地";
            TypeTable.Rows.InsertAt(Typedr, 0);

            this.cmbOgin.DataSource = TypeTable;
            cmbOgin.DisplayMember = "XB001";//显示内容
            //cmbOgin.ValueMember = "XB002";//显示内容

            //CmbDays.ValueMember = "ID";//选项对应的value
            //this.DgvPro1b.DataSource = Myds.Tables[0]; //把第一个表设置为数据源
            //与ASP.NET里GridView控件不同的是这里的DataGridView不需要DataBind()

            if (TypeConnection.State == ConnectionState.Open)
            {
                TypeConnection.Close();
            }
            /// <summary> 
            /// //为CmbDays绑定数据源
            /// </summary>

        }

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
            SqlConnection DTAConnection = new SqlConnection(DTAConnectionString);
            if (e.KeyCode == Keys.Enter)
            {
                if (this.txtINserialno.Text.Length == 0 || this.txtINserialno.Text.Length < 19 || this.cmbOgin.Text == "仕向地" || this.cmbDesc.Text == "目的地")
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
                    if (DTAConnection.State == ConnectionState.Open)
                    {
                        DTAConnection.Close();
                    }
                    txtINserialno.Clear();
                    txtINserialno.Visible = false;
                    TxtMsg.Visible = true;
                    //MessageBox.Show(this.txtproSN.Text.Substring(0, 10) + str);
                    //MessageBox.Show("序号位数不正确，请确认。");
                    //RbtnS.Focus();
                    //if (Convert.ToInt32(CountS.Text) == 0)
                    //{
                    //    CountS.Text = (Convert.ToInt32(CountS.Text)).ToString();
                    //}
                    //else
                    //{
                    //    CountS.Text = (Convert.ToInt32(CountS.Text) - Convert.ToInt32(strsn)).ToString();
                    //}
                    return;
                }

                else
                {

                    DataTable DTAinvTable = new DataTable();
                    string INXBSQL = "SELECT REPLACE(XB001+'-'+XB002,' ','') AS XB001 FROM [LeanSerial].[dbo].[DTASSET_INVXB] WHERE REPLACE(XB002,' ','')='" + txtINserialno.Text.Substring(8, 2) + "' ; ";
                    SqlDataAdapter DTAinvAdapter = new SqlDataAdapter(INXBSQL, DTAConnection);
                    DTAinvAdapter.Fill(DTAinvTable);
                    if (DTAinvTable.Rows[0][0].ToString() != this.cmbOgin.Text)
                    {
                        Sound.Play("error");
                        txtINserialno.Visible = Visible;
                        TxtMsg.Text = "仕向错误，请解锁。正确仕向:'" + DTAinvTable.Rows[0][0].ToString().Trim() + "'" + this.txtINserialno.Text;
                        SqlCommand InsproaSqlcom = new SqlCommand("INSERT INTO [LeanSerial].[dbo].[DTASSET_ERROR] (errorstno,errordate,stuserid,sthostname,sthostip,sthostmac,stmemo)VALUES  ('" + this.txtINserialno.Text + "','" + DateTime.Now.ToString("yyyyMMddHHmmss") + "','AppCube','" + Get_Info.GetComputerName() + "','" + Get_Info.GetIPAddress() + "','" + Get_Info.GetMacAddress() + "','InError') ", DTAConnection);
                        if (DTAConnection.State == ConnectionState.Closed)
                        {
                            DTAConnection.Open();
                        }
                        InsproaSqlcom.ExecuteNonQuery();
                        if (DTAConnection.State == ConnectionState.Open)
                        {
                            DTAConnection.Close();
                        }
                        txtINserialno.Clear();
                        txtINserialno.Visible = false;
                        TxtMsg.Visible = true;
                        //错误播放声音

                        return;
                    }
                    else
                    {


                        try
                        {
                            DataTable DTAStockTable = new DataTable();
                            string DTAStockSQL = "SELECT INSERIAL  FROM [LeanSerial].[dbo].[DTASSET_SCANNER_ALL] WHERE INSERIAL='" + txtINserialno.Text + "' ";
                            SqlDataAdapter DTAStockAdapter = new SqlDataAdapter(DTAStockSQL, DTAConnectionString);
                            DTAStockAdapter.Fill(DTAStockTable);
                            if (DTAStockTable.Rows.Count == 0)
                            {
                                if (DTAConnection.State == ConnectionState.Closed)
                                {
                                    DTAConnection.Open();
                                }
                                Sound.Play("error");
                                txtINserialno.Visible = Visible;
                                TxtMsg.Text = "请先入库再出库，错误需解锁。" + this.txtINserialno.Text;
                                SqlCommand InsproaSqlcom = new SqlCommand("INSERT INTO [LeanSerial].[dbo].[DTASSET_ERROR] (errorstno,errordate,stuserid,sthostname,sthostip,sthostmac,stmemo)VALUES  ('" + this.txtINserialno.Text + "','" + DateTime.Now.ToString("yyyyMMddHHmmss") + "','AppCube','" + Get_Info.GetComputerName() + "','" + Get_Info.GetIPAddress() + "','" + Get_Info.GetMacAddress() + "','InError') ", DTAConnection);
                                InsproaSqlcom.ExecuteNonQuery();
                                if (DTAConnection.State == ConnectionState.Open)
                                {
                                    DTAConnection.Close();
                                }
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

                            }
                            else
                            {


                                DataTable DTAOut = new DataTable();
                                string DTAoutSQL = "SELECT OUTSERIAL  FROM [LeanSerial].[dbo].[DTASSET_SCANNER_ALL] WHERE OUTSERIAL='" + txtINserialno.Text + "' ";
                                SqlDataAdapter DTAoutAdapter = new SqlDataAdapter(DTAoutSQL, DTAConnectionString);
                                DTAoutAdapter.Fill(DTAOut);
                                //显示品名
                                if (DTAOut.Rows.Count == 0)
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
                                    SqlCommand InsproaSqlcom = new SqlCommand("UPDATE [LeanSerial].[dbo].[DTASSET_SCANNER_ALL] SET [OUTINVOICE]='" + this.txtInv.Text + "',[OUTDESC]='" + this.cmbDesc.SelectedValue.ToString() + "',[OUTSERIAL]='" + this.txtINserialno.Text + "',[OUTDATE]='" + dtpDate.Value.ToString("yyyyMMdd") + "',[OUTQTY]=" + strsn + ",[OGION]='" + this.cmbOgin.Text.Substring(0, this.cmbOgin.Text.Trim().Length - 3) + "',[OUTSYSUID]='OutStock',[OUTHNAME]='" + Get_Info.GetComputerName() + "',[OUTHIP]='" + Get_Info.GetIPAddress() + "',[OUTHMAC]='" + Get_Info.GetMacAddress() + "',[OUTUSER]='USER',[OUTDTIME]='" + DateTime.Now.ToString("yyyyMMddHHmmss") + "' WHERE INSERIAL='" + txtINserialno.Text + "'", DTAConnection);
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
                                    TxtMsg.Text = "已经出库，错误需解锁。" + this.txtINserialno.Text;
                                    SqlCommand InsproaSqlcom = new SqlCommand("INSERT INTO [LeanSerial].[dbo].[DTASSET_ERROR] (errorstno,errordate,stuserid,sthostname,sthostip,sthostmac,stmemo)VALUES  ('" + this.txtINserialno.Text + "','" + DateTime.Now.ToString("yyyyMMddHHmmss") + "','AppCube','" + Get_Info.GetComputerName() + "','" + Get_Info.GetIPAddress() + "','" + Get_Info.GetMacAddress() + "','InError') ", DTAConnection);
                                    InsproaSqlcom.ExecuteNonQuery();
                                    if (DTAConnection.State == ConnectionState.Open)
                                    {
                                        DTAConnection.Close();
                                    }
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
        }

        private void TxtMsg_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                if (this.TxtMsg.Text == "OutstockNG")
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
