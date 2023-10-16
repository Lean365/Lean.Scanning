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
namespace Lean.Scanning
{
    public partial class Scanning_Base : UIForm
    {
        public Scanning_Base()
        {
            InitializeComponent();
        }
        public static String DTAConnectionString = ConfigurationManager.ConnectionStrings["SerialDB"].ToString();
        public SqlConnection DTAConnection = new SqlConnection(DTAConnectionString);
        private void button1_Click(object sender, EventArgs e)
        {
            if (this.textBox1.Text == "" || this.textBox2.Text == "")
            {
                Sound.Play("error");
                MessageBox.Show(" 不可空白 ", " 系统提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);
                return;
            }
            else
            {
                DataTable DTAStockTable = new DataTable();
                string DTAStockSQL = "SELECT *  FROM [LeanSerial].[dbo].[DTASSET_INVMB] WHERE MB001='" + this.textBox1.Text.Trim() + "' ";
                SqlDataAdapter DTAStockAdapter = new SqlDataAdapter(DTAStockSQL, DTAConnectionString);
                DTAStockAdapter.Fill(DTAStockTable);
                //显示品名
                if (DTAStockTable.Rows.Count == 0)
                {
                    if (DTAConnection.State == ConnectionState.Closed)
                    {
                        DTAConnection.Open();
                    }
                    SqlCommand InsproaSqlcom = new SqlCommand("INSERT INTO [LeanSerial].[dbo].[DTASSET_INVMB] (MB001,MB002,MB004)VALUES  ('" + this.textBox1.Text.Trim() + "','" + this.textBox2.Text + "','" + this.textBox3.Text + "') ", DTAConnection);
                    InsproaSqlcom.ExecuteNonQuery();
                    textBox1.Text = "";
                    textBox2.Text = "";
                    //textBox3.Text = "";
                    Sound.Play("OK");
                    MessageBox.Show(" 品号追加完成 ", " 系统提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);
                    DataBindmb();
                }
                else
                {
                    Sound.Play("error");
                    MessageBox.Show(" 品号重复 ", " 系统提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);
                    return;
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
        public void DataBindmb()
        {
            string DTAStockSQL = "SELECT [MB001] AS 品号 ,[MB002] AS 品名,[MB004] AS 属性  FROM [LeanSerial].[dbo].[DTASSET_INVMB]";
            SqlDataAdapter da = new SqlDataAdapter(DTAStockSQL, DTAConnectionString);
            da.SelectCommand.CommandTimeout = 0;  //取消超时默认设置  默认是30s   增加一条设置
            DataSet ds = new DataSet();
            da.Fill(ds);
            dataGridView1.DataSource = ds.Tables[0];
        }
        public void DataBindxb()
        {
            string DTAStockSQL = "SELECT [XB001] AS 仕向,[XB002] AS 品号最后两位  FROM [LeanSerial].[dbo].[DTASSET_INVXB]";
            SqlDataAdapter da = new SqlDataAdapter(DTAStockSQL, DTAConnectionString);
            da.SelectCommand.CommandTimeout = 0;  //取消超时默认设置  默认是30s   增加一条设置
            DataSet ds = new DataSet();
            da.Fill(ds);
            dataGridView2.DataSource = ds.Tables[0];
        }
        public void DataBindlb()
        {
            string DTAStockSQL = "SELECT [LB001] AS Destination,[LB002] AS 仕向地  FROM [LeanSerial].[dbo].[DTASSET_INVLB]";
            SqlDataAdapter da = new SqlDataAdapter(DTAStockSQL, DTAConnectionString);
            da.SelectCommand.CommandTimeout = 0;  //取消超时默认设置  默认是30s   增加一条设置
            DataSet ds = new DataSet();
            da.Fill(ds);
            dataGridView3.DataSource = ds.Tables[0];
        }
        private void Scanning_Base_Load(object sender, EventArgs e)
        {
            DataBindmb();
            DataBindxb();
            DataBindlb();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (this.textBox5.Text == "" || this.textBox4.Text == "")
            {
                Sound.Play("error");
                MessageBox.Show(" 不可空白 ", " 系统提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);
                return;
            }
            else
            {

                DataTable DTAStockTable = new DataTable();
                string DTAStockSQL = "SELECT *  FROM [LeanSerial].[dbo].[DTASSET_INVXB] WHERE XB001='" + this.textBox5.Text.Trim() + "' ";
                SqlDataAdapter DTAStockAdapter = new SqlDataAdapter(DTAStockSQL, DTAConnectionString);
                DTAStockAdapter.Fill(DTAStockTable);
                //显示品名
                if (DTAStockTable.Rows.Count == 0)
                {
                    if (DTAConnection.State == ConnectionState.Closed)
                    {
                        DTAConnection.Open();
                    }
                    SqlCommand InsproaSqlcom = new SqlCommand("INSERT INTO [LeanSerial].[dbo].[DTASSET_INVXB] (XB001,XB002)VALUES  ('" + this.textBox5.Text.Trim() + "','" + this.textBox4.Text + "') ", DTAConnection);
                    InsproaSqlcom.ExecuteNonQuery();
                    textBox5.Text = "";
                    textBox4.Text = "";
                    //textBox3.Text = "";
                    Sound.Play("OK");
                    MessageBox.Show(" 品号追加完成 ", " 系统提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);
                    DataBindxb();
                }
                else
                {
                    Sound.Play("error");
                    MessageBox.Show(" 品号重复 ", " 系统提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);
                    return;
                }
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            new System.Threading.Thread((System.Threading.ThreadStart)delegate
            {
                Application.Run(new Scanning_Base());
            }).Start();
            this.Close();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            new System.Threading.Thread((System.Threading.ThreadStart)delegate
            {
                Application.Run(new Scanning_Base());
            }).Start();
            this.Close();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (this.textBox6.Text == "" || this.textBox7.Text == "")
            {
                Sound.Play("error");
                MessageBox.Show(" 不可空白 ", " 系统提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);
                return;
            }
            else
            {

                DataTable DTAStockTable = new DataTable();
                string DTAStockSQL = "SELECT *  FROM [LeanSerial].[dbo].[DTASSET_INVLB] WHERE LB002='" + this.textBox6.Text.Trim() + "' ";
                SqlDataAdapter DTAStockAdapter = new SqlDataAdapter(DTAStockSQL, DTAConnectionString);
                DTAStockAdapter.Fill(DTAStockTable);
                //显示品名
                if (DTAStockTable.Rows.Count == 0)
                {
                    if (DTAConnection.State == ConnectionState.Closed)
                    {
                        DTAConnection.Open();
                    }
                    SqlCommand InsproaSqlcom = new SqlCommand("INSERT INTO [LeanSerial].[dbo].[DTASSET_INVLB] (LB001,LB002)VALUES  ('" + this.textBox7.Text.Trim() + "','" + this.textBox6.Text + "') ", DTAConnection);
                    InsproaSqlcom.ExecuteNonQuery();
                    textBox6.Text = "";
                    textBox7.Text = "";
                    //textBox3.Text = "";
                    Sound.Play("OK");
                    MessageBox.Show(" 目的地追加完成 ", " 系统提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);
                    DataBindlb();
                }
                else
                {
                    Sound.Play("error");
                    MessageBox.Show(" 目的地重复 ", " 系统提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);
                    return;
                }
            }
        }


    }
}
