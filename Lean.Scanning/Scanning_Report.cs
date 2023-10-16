using HFSoft.Component.Windows;
using Sunny.UI;
using System;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Windows.Forms;
namespace Lean.Scanning
{
    public partial class Scanning_Report : UIForm
    {
        public Scanning_Report()
        {
            InitializeComponent();
        }
        public static String DTAConnectionString = ConfigurationManager.ConnectionStrings["SerialDB"].ToString();
        public SqlConnection DTAConnection = new SqlConnection(DTAConnectionString);
        public Int32 IptSerial, strlen;
        public string IptDate, IptHbnSerial, IptHbn, IptStr, Eptinv, Eptorg, EptDesc, strsn;
        [DllImport("kernel32.dll")]
        static extern uint GetTickCount();

        //延时函数
        static void Delay(uint ms)
        {
            uint start = GetTickCount();
            while (GetTickCount() - start < ms)
            {
                Application.DoEvents();
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
        private void btnInstock_Click(object sender, EventArgs e)
        {
            LoadingHandler.Show(this, args =>
            {
                //记时器
                TimeSpan ts1 = new TimeSpan(DateTime.Now.Ticks); //获取当前时间的刻度数

                //延时20分钟1200000
                //Delay(600000);
                Delay(0);
                DataTable DTAStockTable = new DataTable();
                string DTAStockSQL = "DELETE[LeanSerial].[dbo].[DTASSET_SCANNER_IN] WHERE LEFT(IN001,6) = '" + dtpDate.Value.ToString("yyyyMM") + "';SELECT [INSERIAL],[INDATE],[INQTY]  FROM[LeanSerial].[dbo].[DTASSET_SCANNER_ALL] WHERE LEFT(INDATE,6)='" + dtpDate.Value.ToString("yyyyMM") + "'";
                SqlDataAdapter DTAStockAdapter = new SqlDataAdapter(DTAStockSQL, DTAConnectionString);
                DTAStockAdapter.Fill(DTAStockTable);
                //显示品名
                if (DTAStockTable.Rows.Count >= 1)
                {

                    for (int tbi = 1; tbi <= DTAStockTable.Rows.Count; tbi++)
                    {
                        if (DTAStockTable.Rows[tbi - 1][0].ToString().Length != 0)
                        {
                            IptHbnSerial = DTAStockTable.Rows[tbi - 1][0].ToString();
                            IptHbn = DTAStockTable.Rows[tbi - 1][0].ToString().Substring(0, 10);
                            IptDate = DTAStockTable.Rows[tbi - 1][1].ToString();


                            int fs = Convert.ToInt32(DTAStockTable.Rows[tbi - 1][0].ToString().Length);
                            int tsss = Convert.ToInt32(DTAStockTable.Rows[tbi - 1][0].ToString().Substring(fs - 2, 2));
                            if (tsss == 0)
                            {
                                IptSerial = Convert.ToInt32(DTAStockTable.Rows[tbi - 1][0].ToString().Substring(fs - 3, 3));
                                for (int i = 0; i < IptSerial; i++)
                                {
                                    string strA = DTAStockTable.Rows[tbi - 1][0].ToString().Substring(10, fs - 10 - 3);
                                    Regex reg = new Regex("^[0-9]+$");
                                    Match ma = reg.Match(strA);
                                    /// <summary>
                                    /// //判断是数字字串还是字符字串
                                    /// </summary>

                                    if (ma.Success)
                                    {
                                        int len = Convert.ToInt32(DTAStockTable.Rows[tbi - 1][0].ToString().Substring(10, fs - 10 - 3)) + i;
                                        if (Convert.ToInt32(len.ToString().Length) != 7)
                                        {
                                            IptStr = String.Format("{0:D7}", len);
                                        }
                                        else
                                        {
                                            IptStr = len.ToString();
                                        }
                                        Iptsave();
                                    }
                                    else
                                    {
                                        //判断字母出现的位置
                                        //int f = 0;
                                        //foreach (char c in DTAStockTable.Rows[tbi - 1][0].ToString().Substring(10, 7))
                                        //{
                                        //    f++;
                                        //    if (char.IsLetter(c))
                                        //    {
                                        //        //int iii = ii - 1;
                                        //        string ss = DTAStockTable.Rows[tbi - 1][0].ToString().Substring(10, 7).Substring(f, 7 - f);
                                        //        string sss = DTAStockTable.Rows[tbi - 1][0].ToString().Substring(10, 7).Substring(0, f);
                                        //    }
                                        //}
                                        char cc = DTAStockTable.Rows[tbi - 1][0].ToString().Substring(10, fs - 10 - 3).Last((c) => { return (c >= 'a' && c <= 'z') || (c >= 'A' && c <= 'Z'); });
                                        int index = DTAStockTable.Rows[tbi - 1][0].ToString().Substring(10, fs - 10 - 3).LastIndexOf(cc) + 1;
                                        string sss = DTAStockTable.Rows[tbi - 1][0].ToString().Substring(10, fs - 10 - 3).Substring(0, index);
                                        string ss = DTAStockTable.Rows[tbi - 1][0].ToString().Substring(10, fs - 10 - 3).Substring(index, fs - 10 - 3 - index);

                                        Int64 len = Convert.ToInt64(ss) + i;
                                        Int64 lena = Convert.ToInt64(ss.Length.ToString());
                                        if (Convert.ToInt32(len.ToString().Length) != lena)
                                        {
                                            string d = "{0:D" + lena + "}";
                                            IptStr = String.Format(d, len);
                                            //str = len.ToString(d);
                                        }
                                        else
                                        {
                                            IptStr = len.ToString();
                                        }
                                        //判断字母出现的位置

                                        IptStr = sss + Convert.ToString(IptStr);
                                        /// <summary>
                                        /// //判断是数字字串还是字符字串
                                        /// </summary>
                                        Iptsave();
                                    }
                                }
                            }
                            else
                            {
                                IptSerial = Convert.ToInt32(DTAStockTable.Rows[tbi - 1][0].ToString().Substring(fs - 2, 2));
                                for (int i = 0; i < IptSerial; i++)
                                {
                                    string strA = DTAStockTable.Rows[tbi - 1][0].ToString().Substring(10, fs - 10 - 2);
                                    Regex reg = new Regex("^[0-9]+$");
                                    Match ma = reg.Match(strA);
                                    /// <summary>
                                    /// //判断是数字字串还是字符字串
                                    /// </summary>

                                    if (ma.Success)
                                    {
                                        int len = Convert.ToInt32(DTAStockTable.Rows[tbi - 1][0].ToString().Substring(10, fs - 10 - 2)) + i;
                                        if (Convert.ToInt32(len.ToString().Length) != 7)
                                        {
                                            IptStr = String.Format("{0:D7}", len);
                                        }
                                        else
                                        {
                                            IptStr = len.ToString();
                                        }
                                        Iptsave();
                                    }
                                    else
                                    {
                                        //判断字母出现的位置
                                        //int f = 0;
                                        //foreach (char c in DTAStockTable.Rows[tbi - 1][0].ToString().Substring(10, 7))
                                        //{
                                        //    f++;
                                        //    if (char.IsLetter(c))
                                        //    {
                                        //        //int iii = ii - 1;
                                        //        string ss = DTAStockTable.Rows[tbi - 1][0].ToString().Substring(10, 7).Substring(f, 7 - f);
                                        //        string sss = DTAStockTable.Rows[tbi - 1][0].ToString().Substring(10, 7).Substring(0, f);
                                        //    }
                                        //}
                                        char cc = DTAStockTable.Rows[tbi - 1][0].ToString().Substring(10, fs - 10 - 2).Last((c) => { return (c >= 'a' && c <= 'z') || (c >= 'A' && c <= 'Z'); });
                                        int index = DTAStockTable.Rows[tbi - 1][0].ToString().Substring(10, fs - 10 - 2).LastIndexOf(cc) + 1;
                                        string sss = DTAStockTable.Rows[tbi - 1][0].ToString().Substring(10, fs - 10 - 2).Substring(0, index);
                                        string ss = DTAStockTable.Rows[tbi - 1][0].ToString().Substring(10, fs - 10 - 2).Substring(index, fs - 10 - 2 - index);

                                        Int64 len = Convert.ToInt64(ss) + i;
                                        Int64 lena = Convert.ToInt64(ss.Length.ToString());
                                        if (Convert.ToInt32(len.ToString().Length) != lena)
                                        {
                                            string d = "{0:D" + lena + "}";
                                            IptStr = String.Format(d, len);
                                            //str = len.ToString(d);
                                        }
                                        else
                                        {
                                            IptStr = len.ToString();
                                        }
                                        //判断字母出现的位置

                                        IptStr = sss + Convert.ToString(IptStr);
                                        /// <summary>
                                        /// //判断是数字字串还是字符字串
                                        /// </summary>
                                        Iptsave();
                                    }
                                }
                            }

                            Sound.Play("ok");

                        }
                        else
                        {
                            Sound.Play("error");
                            return;


                        }
                    }
                }
                else
                {
                    Sound.Play("error");
                    return;

                }
                IptProcess();

                MessageBox.Show(" 入库报表处理完成 ", " 系统提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);
            });
        }

        private void button2_Click(object sender, EventArgs e)
        {
            DataTable DTAStockTable = new DataTable();
            string DTAStockSQL = "SELECT      [INS001] AS InDate,[INS002] AS ItemMaster,[INS003] AS SerialNumber,[INS004] AS Qty,[MB002] AS ItemText " +
                " FROM[LeanSerial].[dbo].[DTASSET_SCANNER_IN_SUB]  LEFT JOIN[LeanSerial].[dbo].[DTASSET_INVMB]  ON MB001 = INS002 WHERE LEFT(INS001,6)='" + dtpDate.Value.ToString("yyyyMM") + "'";
            SqlDataAdapter DTAStockAdapter = new SqlDataAdapter(DTAStockSQL, DTAConnectionString);
            DTAStockAdapter.Fill(DTAStockTable);
            //显示品名
            if (DTAStockTable.Rows.Count >= 1)
            {
                string path = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
                string filen = path + "\\" + dtpDate.Value.ToString("yyyyMM") + "_入库扫描明细.xlsx";

                Excel_Npoi.TableToExcelForXLSX(DTAStockTable, filen);

                MessageBox.Show(filen + " 导出完成 ", " 系统提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            DataTable DTAStockTable = new DataTable();
            string DTAStockSQL = "SELECT[OUTS002] AS InvoiceNo,[OUTS001] AS ShipDate,[OUTS004] AS ItemMaster,[OUTS005] AS SerialNumber,[OUTS006] AS Qty,[MB002] AS ItemText,OUTS008 AS RecipientName " +
                                 "FROM[LeanSerial].[dbo].[DTASSET_SCANNER_OUT_SUB]LEFT JOIN[LeanSerial].[dbo].[DTASSET_INVMB]  ON MB001 =[OUTS004] WHERE LEFT(OUTS001,6)='" + dtpDate.Value.ToString("yyyyMM") + "'";
            SqlDataAdapter DTAStockAdapter = new SqlDataAdapter(DTAStockSQL, DTAConnectionString);
            DTAStockAdapter.Fill(DTAStockTable);
            //显示品名
            if (DTAStockTable.Rows.Count >= 1)
            {
                string path = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
                string filen = path + "\\" + dtpDate.Value.ToString("yyyyMM") + "_出货扫描明细_月报.xlsx";

                Excel_Npoi.TableToExcelForXLSX(DTAStockTable, filen);
                MessageBox.Show(filen + " 导出完成 ", " 系统提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);
            }
        }
        public void DataBindALL()
        {
            string DTAStockSQL = "SELECT [INSERIAL] AS InSerial ,[INDATE] AS InDate ,[INQTY]AS InQty " +
                                    " ,[OUTINVOICE]AS Invoice,[OUTDESC]AS Transport ,[OUTSERIAL]AS OutkSerial,[OUTDATE]AS OutDate " +
                                    " ,[OUTQTY]AS Outqty,[OGION]AS Destination  FROM[LeanSerial].[dbo].[DTASSET_SCANNER_ALL] WHERE LEFT(INDATE,6) = '" + dtpDate.Value.ToString("yyyyMM") + "'";
            SqlDataAdapter da = new SqlDataAdapter(DTAStockSQL, DTAConnectionString);
            //重要的在这个地方,这里设置为0的法,就永远不会出现超时;
            da.SelectCommand.CommandTimeout = 0;  //取消超时默认设置  默认是30s   增加一条设置
            DataSet ds = new DataSet();
            da.Fill(ds);
            dataGridView1.DataSource = ds.Tables[0];
        }

        private void button3_Click(object sender, EventArgs e)
        {
            DataTable DTAStockTable = new DataTable();
            string DTAStockSQL = "SELECT[OUTS002] AS InvoiceNo,[OUTS001] AS ShipDate,[OUTS004] AS ItemMaster,[OUTS005] AS SerialNumber,[OUTS006] AS Qty,[MB002] AS ItemText,OUTS008 AS RecipientName " +
                                 "FROM[LeanSerial].[dbo].[DTASSET_SCANNER_OUT_SUB]LEFT JOIN[LeanSerial].[dbo].[DTASSET_INVMB]  ON MB001 =[OUTS004] WHERE OUTS001='" + dtpDate.Value.ToString("yyyyMMdd") + "'";
            SqlDataAdapter DTAStockAdapter = new SqlDataAdapter(DTAStockSQL, DTAConnectionString);
            DTAStockAdapter.Fill(DTAStockTable);
            //显示品名
            if (DTAStockTable.Rows.Count >= 1)
            {
                string path = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
                string filen = path + "\\" + dtpDate.Value.ToString("yyyyMMdd") + "_出货扫描明细_日报.xlsx";

                Excel_Npoi.TableToExcelForXLSX(DTAStockTable, filen);
                MessageBox.Show(filen + " 导出完成 ", " 系统提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);
            }
        }

        public void DataBindIN()
        {
            string DTAStockSQL = "SELECT [INS001] AS StockDate ,[INS002] AS Item ,[INS003]AS StockSerial " +
                                    " ,[INS004]AS StockQty,[INS005]AS SerialAll,MB002 AS ItemText " +
                                    "   FROM[LeanSerial].[dbo].[DTASSET_SCANNER_IN_SUB] left join [LeanSerial].[dbo].[DTASSET_INVMB] on MB001=INS002";
            SqlDataAdapter da = new SqlDataAdapter(DTAStockSQL, DTAConnectionString);
            DataSet ds = new DataSet();
            da.Fill(ds);
            dataGridView1.DataSource = ds.Tables[0];
        }
        public void DataBindOUT()
        {
            string DTAStockSQL = "SELECT [OUTS002] AS InvoiceNo ,[OUTS001] AS ShipDate ,[OUTS006]AS Qty " +
                                    " ,[OUTS004]AS Item,[OUTS005]AS SerialNo,[OUTS007]AS SerialNoAll MB002 AS ItemText,OUTS003 AS Destination " +
                                    " FROM [LeanSerial].[dbo].[DTASSET_SCANNER_OUT_SUB] left join [LeanSerial].[dbo].[DTASSET_INVMB] on MB001=OUTS004;";
            SqlDataAdapter da = new SqlDataAdapter(DTAStockSQL, DTAConnectionString);
            DataSet ds = new DataSet();
            da.Fill(ds);
            dataGridView1.DataSource = ds.Tables[0];
        }

        private void fmProcess_Load(object sender, EventArgs e)
        {
            DataBindALL();
        }

        public void Iptsave()
        {
            if (DTAConnection.State == ConnectionState.Closed)
            {
                DTAConnection.Open();
            }
            SqlCommand InsproaSqlcom = new SqlCommand("INSERT INTO [LeanSerial].[dbo].[DTASSET_SCANNER_IN] ([IN001],[IN002],[IN003],[IN004],[IN005],[IN006],[INSYSUID],[INHNAME],[INHIP],[INHMAC],[INUSER],[INDTIME],[studfchar1],[studfchar2],[studfchar3],[studfchar4],[studfint1],[studfint2],[studfint3],[studfint4])VALUES  ('" + IptDate + "','" + this.IptHbnSerial + "','" + this.IptHbn + "', '" + this.IptStr + "','',1,'InStock','" + Get_Info.GetComputerName() + "','" + Get_Info.GetIPAddress() + "','" + Get_Info.GetMacAddress() + "','USER','" + DateTime.Now.ToString("yyyyMMddHHmmss") + "','','','','','0','0','0','0') ", DTAConnection);
            InsproaSqlcom.ExecuteNonQuery();
            if (DTAConnection.State == ConnectionState.Open)
            {
                DTAConnection.Close();
            }
        }
        public void IptProcess()
        {
            DataTable dtShip = new DataTable();
            string dtShipSQL = "SELECT *  FROM [LeanSerial].[dbo].[DTASSET_SCANNER_IN] WHERE LEFT(IN001,6)='" + dtpDate.Value.ToString("yyyyMM") + "'";
            SqlDataAdapter daShip = new SqlDataAdapter(dtShipSQL, DTAConnectionString);
            daShip.Fill(dtShip);
            //显示品名
            if (dtShip.Rows.Count >= 1)
            {
                DateTime strdate = DateTime.Now;
                //string struser = GetIdentityName();

                SqlCommand Iptrep = new SqlCommand("DELETE[LeanSerial].[dbo].[DTASSET_SCANNER_IN_SUB] WHERE LEFT(INS001,6)='" + dtpDate.Value.ToString("yyyyMM") + "'; insert[LeanSerial].[dbo].[DTASSET_SCANNER_IN_SUB]([INS001],[INS002],[INS003],[INS004],[INS005],[INSYSUID],[INHNAME],[INDTIME]) select IN001, IN003 as IN002," +
                    " lIN004+right(cast(power(10,4) as varchar )+min(convert(int, IN004)),4) +'~'+lIN004+right(cast(power(10,4) as varchar)+max(convert(int, IN004)),4) as IN003,max(convert(int, IN004))-min(convert(int, IN004))+1 as IN004,IN002 AS INS005,'" +
                    " InReport' as [INSYSUID],'" + Get_Info.GetComputerName() + "' as [INHNAME],'" + strdate.ToString("yyyyMMddhhss") + "'as [INDTIME] from " +
                    " (select row_number()over(order by IN004) as rid, left(IN001, 8) as IN001, IN003, right(IN004, 4) as IN004, " +
                                              " case when len(IN004) = 50 then LEFT(IN004, 46) " +
                                              " when len(IN004) = 59 then LEFT(IN004, 45) " +
                                              " when len(IN004) = 48 then LEFT(IN004, 44) " +
                                              " when len(IN004) = 47 then LEFT(IN004, 43) " +
                                              " when len(IN004) = 46 then LEFT(IN004, 42) " +
                                              " when len(IN004) = 45 then LEFT(IN004, 41) " +
                                              " when len(IN004) = 44 then LEFT(IN004, 40) " +
                                              " when len(IN004) = 43 then LEFT(IN004, 39) " +
                                              " when len(IN004) = 42 then LEFT(IN004, 38) " +
                                              " when len(IN004) = 41 then LEFT(IN004, 37) " +
                                              " when len(IN004) = 40 then LEFT(IN004, 36) " +
                                              " when len(IN004) = 39 then LEFT(IN004, 35) " +
                                              " when len(IN004) = 38 then LEFT(IN004, 34) " +
                                              " when len(IN004) = 37 then LEFT(IN004, 33) " +
                                              " when len(IN004) = 36 then LEFT(IN004, 32) " +
                                              " when len(IN004) = 35 then LEFT(IN004, 31) " +
                                              " when len(IN004) = 34 then LEFT(IN004, 30) " +
                                              " when len(IN004) = 33 then LEFT(IN004, 29) " +
                                              " when len(IN004) = 32 then LEFT(IN004, 28) " +
                                              " when len(IN004) = 31 then LEFT(IN004, 27) " +
                                              " when len(IN004) = 30 then LEFT(IN004, 26) " +
                                              " when len(IN004) = 29 then LEFT(IN004, 25) " +
                                              " when len(IN004) = 28 then LEFT(IN004, 24) " +
                                              " when len(IN004) = 27 then LEFT(IN004, 23) " +
                                              " when len(IN004) = 26 then LEFT(IN004, 22) " +
                                              " when len(IN004) = 25 then LEFT(IN004, 21) " +
                                              " when len(IN004) = 24 then LEFT(IN004, 20) " +
                                              " when len(IN004) = 23 then LEFT(IN004, 19) " +
                                              " when len(IN004) = 22 then LEFT(IN004, 18) " +
                                              " when len(IN004) = 21 then LEFT(IN004, 17) " +
                                              " when len(IN004) = 20 then LEFT(IN004, 16) " +
                                              " when len(IN004) = 19 then LEFT(IN004, 15) " +
                                              " when len(IN004) = 18 then LEFT(IN004, 14) " +
                                              " when len(IN004) = 17 then LEFT(IN004, 13) " +
                                              " when len(IN004) = 16 then LEFT(IN004, 12) " +
                                              " when len(IN004) = 15 then LEFT(IN004, 11) " +
                                              " when len(IN004) = 14 then LEFT(IN004, 10) " +
                                              " when len(IN004) = 13 then LEFT(IN004, 9) " +
                                              " when len(IN004) = 12 then LEFT(IN004, 8) " +
                                              " when len(IN004) = 11 then LEFT(IN004, 7) " +
                                              " when len(IN004) = 10 then LEFT(IN004, 6) " +
                                              " when len(IN004) = 9 then LEFT(IN004, 5) " +
                                              " when len(IN004) = 8 then LEFT(IN004, 4) " +
                                              " when len(IN004) = 7 then LEFT(IN004, 3) " +
                                              " else left(IN004, 4) end as lIN004,IN002 " +
                        " from[dbo].[DTASSET_SCANNER_IN] where LEN(IN004) >= 7 AND LEFT(IN001,6)='" + dtpDate.Value.ToString("yyyyMM") + "'  " +
                        "                      )SHIP group by IN001,IN002, lIN004, IN003, convert(int, IN004) - convert(int, rid) having(count(1) > 0) order by IN001; ", DTAConnection);
                if (DTAConnection.State != ConnectionState.Closed)
                {
                    DTAConnection.Close();
                }

                DTAConnection.Open();
                Iptrep.ExecuteNonQuery();
                DTAConnection.Close();
                Sound.Play("ok");
                //DataBindIN();
            }
            else
            {
                Sound.Play("error");
                return;
            }




        }
        public void Eptsave()
        {
            if (DTAConnection.State == ConnectionState.Closed)
            {
                DTAConnection.Open();
            }
            SqlCommand InsproaSqlcom = new SqlCommand("INSERT INTO [LeanSerial].[dbo].[DTASSET_SCANNER_OUT] ([OUTS001],[OUTS002],[OUTS003],[OUTS004],[OUTS005],[OUTS006],[OUTS007],[OUTS008],[OUTSYSUID],[OUTHNAME],[OUTHIP],[OUTHMAC],[OUTUSER],[OUTDTIME],[studfchar1],[studfchar2],[studfchar3],[studfchar4],[studfint1],[studfint2],[studfint3],[studfint4])VALUES  ('" + IptDate + "','" + this.Eptinv + "','" + this.Eptorg + "', '" + this.IptHbnSerial + "','" + this.IptHbn + "','" + this.IptStr + "','" + this.EptDesc + "',1,'OutStock','" + Get_Info.GetComputerName() + "','" + Get_Info.GetIPAddress() + "','" + Get_Info.GetMacAddress() + "','USER','" + DateTime.Now.ToString("yyyyMMddHHmmss") + "','','','','','0','0','0','0') ", DTAConnection);
            InsproaSqlcom.ExecuteNonQuery();
            if (DTAConnection.State == ConnectionState.Open)
            {
                DTAConnection.Close();
            }
        }
        public void EptProcess()
        {
            DataTable dtShip = new DataTable();
            string dtShipSQL = "SELECT *  FROM [LeanSerial].[dbo].[DTASSET_SCANNER_OUT] WHERE LEFT(OUTS001,6)='" + dtpDate.Value.ToString("yyyyMM") + "'";
            SqlDataAdapter daShip = new SqlDataAdapter(dtShipSQL, DTAConnection);
            daShip.Fill(dtShip);
            //显示品名
            if (dtShip.Rows.Count >= 1)
            {
                DateTime strdate = DateTime.Now;

                //string struser = GetIdentityName();

                SqlCommand Eptrep = new SqlCommand(" DELETE[LeanSerial].[dbo].[DTASSET_SCANNER_OUT_SUB] WHERE LEFT(OUTS001,6)='" + dtpDate.Value.ToString("yyyyMM") + "'; " +
                " insert [LeanSerial].[dbo].[DTASSET_SCANNER_OUT_SUB](OUTS001,OUTS002,OUTS003,OUTS004,OUTS005,OUTS006,OUTS007,OUTS008,[OUTSYSUID],[OUTHNAME],[OUTDTIME])   " +
                " select  OUTS001, OUTS002, OUTS003, OUTS005 as OUTS004, lOUTS006 + right(cast(power(10, 4) as varchar) + min(convert(int, OUTS006)), 4) + '~' + " +
                                        " lOUTS006 + right(cast(power(10, 4) as varchar) + max(convert(int, OUTS006)), 4) as OUTS005,  max(convert(int, OUTS006)) - min(convert(int, OUTS006)) + 1 as OUTS006,						" +
                                        " OUTS004 AS OUTS007,OUTS007 AS OUTS008,'OutReport'[OUTSYSUID],'" + Get_Info.GetComputerName() + "'[OUTHNAME],'" + strdate.ToString("yyyyMMddhhss") + "'[OUTDTIME] from   (						" +
                                        " select row_number()over(order by OUTS004) as rid, left(OUTS001, 8) as OUTS001,   OUTS002, OUTS003, OUTS005, right(OUTS006, 4) as OUTS006,   " +
                                        " case when len(OUTS006) = 50    then LEFT(OUTS006, 46)   " +
                                        " when len(OUTS006) = 49   then LEFT(OUTS006, 45)   " +
                                        " when len(OUTS006) = 48   then LEFT(OUTS006, 44)   " +
                                        " when len(OUTS006) = 47   then LEFT(OUTS006, 43)   " +
                                        " when len(OUTS006) = 46   then LEFT(OUTS006, 42)   " +
                                        " when len(OUTS006) = 45   then LEFT(OUTS006, 41)   " +
                                        " when len(OUTS006) = 44   then LEFT(OUTS006, 40)   " +
                                        " when len(OUTS006) = 43   then LEFT(OUTS006, 39)   " +
                                        " when len(OUTS006) = 42   then LEFT(OUTS006, 38)   " +
                                        " when len(OUTS006) = 41   then LEFT(OUTS006, 37)   " +
                                        " when len(OUTS006) = 40   then LEFT(OUTS006, 36)   " +
                                        " when len(OUTS006) = 39   then LEFT(OUTS006, 35)   " +
                                        " when len(OUTS006) = 38   then LEFT(OUTS006, 34)   " +
                                        " when len(OUTS006) = 37   then LEFT(OUTS006, 33)   " +
                                        " when len(OUTS006) = 36   then LEFT(OUTS006, 32)   " +
                                        " when len(OUTS006) = 35   then LEFT(OUTS006, 31)   " +
                                        " when len(OUTS006) = 34   then LEFT(OUTS006, 30)   " +
                                        " when len(OUTS006) = 33   then LEFT(OUTS006, 29)   " +
                                        " when len(OUTS006) = 32   then LEFT(OUTS006, 28)   " +
                                        " when len(OUTS006) = 31   then LEFT(OUTS006, 27)   " +
                                        " when len(OUTS006) = 30   then LEFT(OUTS006, 26)   " +
                                        " when len(OUTS006) = 29   then LEFT(OUTS006, 25)   " +
                                        " when len(OUTS006) = 28   then LEFT(OUTS006, 24)   " +
                                        " when len(OUTS006) = 27   then LEFT(OUTS006, 23)   " +
                                        " when len(OUTS006) = 26   then LEFT(OUTS006, 22)   " +
                                        " when len(OUTS006) = 25 then LEFT(OUTS006, 21)   " +
                                        " when len(OUTS006) = 24 then LEFT(OUTS006, 20)   " +
                                        " when len(OUTS006) = 23 then LEFT(OUTS006, 19)   " +
                                        " when len(OUTS006) = 22 then LEFT(OUTS006, 18)   " +
                                        " when len(OUTS006) = 21 then LEFT(OUTS006, 17)   " +
                                        " when len(OUTS006) = 20 then LEFT(OUTS006, 16)   " +
                                        " when len(OUTS006) = 19 then LEFT(OUTS006, 15)   " +
                                        " when len(OUTS006) = 18 then LEFT(OUTS006, 14)   " +
                                        " when len(OUTS006) = 17 then LEFT(OUTS006, 13)   " +
                                        " when len(OUTS006) = 16 then LEFT(OUTS006, 12)   " +
                                        " when len(OUTS006) = 15 then LEFT(OUTS006, 11)   " +
                                        " when len(OUTS006) = 14 then LEFT(OUTS006, 10)   " +
                                        " when len(OUTS006) = 13 then LEFT(OUTS006, 9)   " +
                                        " when len(OUTS006) = 12 then LEFT(OUTS006, 8)   " +
                                        " when len(OUTS006) = 11 then LEFT(OUTS006, 7)   " +
                                        " when len(OUTS006) = 10 then LEFT(OUTS006, 6)   " +
                                        " when len(OUTS006) = 9 then LEFT(OUTS006, 5)   " +
                                        " when len(OUTS006) = 8 then LEFT(OUTS006, 4)   " +
                                        " when len(OUTS006) = 7 then LEFT(OUTS006, 3)      " +
                                        " else left(OUTS006, 4) end as lOUTS006,						" +
                                        " OUTS004,OUTS007  " +
                                        " from[LeanSerial].[dbo].[DTASSET_SCANNER_OUT] where  LEN(OUTS006) >= 7  AND LEFT(OUTS001, 6) = '" + dtpDate.Value.ToString("yyyyMM") + "'     " +
                                        " )SHIP group by OUTS001,OUTS004,OUTS007, OUTS002, OUTS003, OUTS005, lOUTS006, " +
                                        " convert(int, OUTS006) - convert(int, rid) having(count(1) > 0) order by OUTS001;", DTAConnection);
                if (DTAConnection.State != ConnectionState.Closed)
                {
                    DTAConnection.Close();
                }
                DTAConnection.Open();
                Eptrep.ExecuteNonQuery();
                DTAConnection.Close();
                Sound.Play("ok");
                //DataBindOUT();
            }
            else
            {
                Sound.Play("error");
                return;
            }
        }
        private void btnOutstock_Click(object sender, EventArgs e)
        {
            LoadingHandler.Show(this, args =>
            {
                //记时器
                TimeSpan ts1 = new TimeSpan(DateTime.Now.Ticks); //获取当前时间的刻度数

                //延时20分钟1200000
                //Delay(600000);
                Delay(0);
                DataTable DTAStockTable = new DataTable();
                string DTAStockSQL = "DELETE[LeanSerial].[dbo].[DTASSET_SCANNER_OUT] WHERE LEFT(OUTS001,6) = '" + dtpDate.Value.ToString("yyyyMM") + "';SELECT [OUTSERIAL], [OUTINVOICE],[OUTDESC],[OUTDATE],[OUTQTY],[OGION]  FROM [LeanSerial].[dbo].[DTASSET_SCANNER_ALL] WHERE LEFT([OUTDATE],6)='" + dtpDate.Value.ToString("yyyyMM") + "'";
                SqlDataAdapter DTAStockAdapter = new SqlDataAdapter(DTAStockSQL, DTAConnectionString);
                DTAStockAdapter.Fill(DTAStockTable);
                //显示品名
                if (DTAStockTable.Rows.Count >= 1)
                {

                    for (int tbi = 1; tbi <= DTAStockTable.Rows.Count; tbi++)
                    {
                        if (DTAStockTable.Rows[tbi - 1][0].ToString().Length != 0)
                        {
                            int fs = Convert.ToInt32(DTAStockTable.Rows[tbi - 1][0].ToString().Length);
                            int tsss = Convert.ToInt32(DTAStockTable.Rows[tbi - 1][0].ToString().Substring(fs - 2, 2));

                            IptSerial = tsss;
                            IptHbnSerial = DTAStockTable.Rows[tbi - 1][0].ToString();
                            IptHbn = DTAStockTable.Rows[tbi - 1][0].ToString().Substring(0, 10);
                            IptDate = DTAStockTable.Rows[tbi - 1][3].ToString();
                            Eptinv = DTAStockTable.Rows[tbi - 1][1].ToString();
                            Eptorg = DTAStockTable.Rows[tbi - 1][5].ToString();
                            EptDesc = DTAStockTable.Rows[tbi - 1][2].ToString();

                            if (IptSerial == 0)
                            {
                                IptSerial = Convert.ToInt32(DTAStockTable.Rows[tbi - 1][0].ToString().Substring(fs - 3, 3));
                                for (int i = 0; i < IptSerial; i++)
                                {
                                    string strA = DTAStockTable.Rows[tbi - 1][0].ToString().Substring(10, fs - 10 - 3);
                                    Regex reg = new Regex("^[0-9]+$");
                                    Match ma = reg.Match(strA);
                                    /// <summary>
                                    /// //判断是数字字串还是字符字串
                                    /// </summary>

                                    if (ma.Success)
                                    {
                                        int len = Convert.ToInt32(DTAStockTable.Rows[tbi - 1][0].ToString().Substring(10, fs - 10 - 3)) + i;
                                        if (Convert.ToInt32(len.ToString().Length) != 7)
                                        {
                                            IptStr = String.Format("{0:D7}", len);
                                        }
                                        else
                                        {
                                            IptStr = len.ToString();
                                        }
                                        Eptsave();
                                    }
                                    else
                                    {
                                        //判断字母出现的位置
                                        //int f = 0;
                                        //foreach (char c in DTAStockTable.Rows[tbi - 1][0].ToString().Substring(10, 7))
                                        //{
                                        //    f++;
                                        //    if (char.IsLetter(c))
                                        //    {
                                        //        //int iii = ii - 1;
                                        //        string ss = DTAStockTable.Rows[tbi - 1][0].ToString().Substring(10, 7).Substring(f, 7 - f);
                                        //        string sss = DTAStockTable.Rows[tbi - 1][0].ToString().Substring(10, 7).Substring(0, f);
                                        //    }
                                        //}
                                        char cc = DTAStockTable.Rows[tbi - 1][0].ToString().Substring(10, fs - 10 - 3).Last((c) => { return (c >= 'a' && c <= 'z') || (c >= 'A' && c <= 'Z'); });
                                        int index = DTAStockTable.Rows[tbi - 1][0].ToString().Substring(10, fs - 10 - 3).LastIndexOf(cc) + 1;
                                        string sss = DTAStockTable.Rows[tbi - 1][0].ToString().Substring(10, fs - 10 - 3).Substring(0, index);
                                        string ss = DTAStockTable.Rows[tbi - 1][0].ToString().Substring(10, fs - 10 - 3).Substring(index, fs - 10 - 3 - index);

                                        int len = Convert.ToInt32(ss) + i;
                                        int lena = Convert.ToInt32(ss.Length.ToString());
                                        if (Convert.ToInt32(len.ToString().Length) != lena)
                                        {
                                            string d = "{0:D" + lena + "}";
                                            IptStr = String.Format(d, len);
                                            //str = len.ToString(d);
                                        }
                                        else
                                        {
                                            IptStr = len.ToString();
                                        }
                                        //判断字母出现的位置

                                        IptStr = sss + Convert.ToString(IptStr);
                                        /// <summary>
                                        /// //判断是数字字串还是字符字串
                                        /// </summary>
                                        Eptsave();
                                    }
                                }
                            }


                            else
                            {
                                IptSerial = Convert.ToInt32(DTAStockTable.Rows[tbi - 1][0].ToString().Substring(fs - 2, 2));
                                for (int i = 0; i < IptSerial; i++)
                                {
                                    string strA = DTAStockTable.Rows[tbi - 1][0].ToString().Substring(10, fs - 10 - 2);
                                    Regex reg = new Regex("^[0-9]+$");
                                    Match ma = reg.Match(strA);
                                    /// <summary>
                                    /// //判断是数字字串还是字符字串
                                    /// </summary>

                                    if (ma.Success)
                                    {
                                        int len = Convert.ToInt32(DTAStockTable.Rows[tbi - 1][0].ToString().Substring(10, fs - 10 - 2)) + i;
                                        if (Convert.ToInt32(len.ToString().Length) != 7)
                                        {
                                            IptStr = String.Format("{0:D7}", len);
                                        }
                                        else
                                        {
                                            IptStr = len.ToString();
                                        }
                                        Eptsave();
                                    }
                                    else
                                    {
                                        //判断字母出现的位置
                                        //int f = 0;
                                        //foreach (char c in DTAStockTable.Rows[tbi - 1][0].ToString().Substring(10, 7))
                                        //{
                                        //    f++;
                                        //    if (char.IsLetter(c))
                                        //    {
                                        //        //int iii = ii - 1;
                                        //        string ss = DTAStockTable.Rows[tbi - 1][0].ToString().Substring(10, 7).Substring(f, 7 - f);
                                        //        string sss = DTAStockTable.Rows[tbi - 1][0].ToString().Substring(10, 7).Substring(0, f);
                                        //    }
                                        //}
                                        char cc = DTAStockTable.Rows[tbi - 1][0].ToString().Substring(10, fs - 10 - 2).Last((c) => { return (c >= 'a' && c <= 'z') || (c >= 'A' && c <= 'Z'); });
                                        int index = DTAStockTable.Rows[tbi - 1][0].ToString().Substring(10, fs - 10 - 2).LastIndexOf(cc) + 1;
                                        string sss = DTAStockTable.Rows[tbi - 1][0].ToString().Substring(10, fs - 10 - 2).Substring(0, index);
                                        string ss = DTAStockTable.Rows[tbi - 1][0].ToString().Substring(10, fs - 10 - 2).Substring(index, fs - 10 - 2 - index);

                                        int len = Convert.ToInt32(ss) + i;
                                        int lena = Convert.ToInt32(ss.Length.ToString());
                                        if (Convert.ToInt32(len.ToString().Length) != lena)
                                        {
                                            string d = "{0:D" + lena + "}";
                                            IptStr = String.Format(d, len);
                                            //str = len.ToString(d);
                                        }
                                        else
                                        {
                                            IptStr = len.ToString();
                                        }
                                        //判断字母出现的位置

                                        IptStr = sss + Convert.ToString(IptStr);
                                        /// <summary>
                                        /// //判断是数字字串还是字符字串
                                        /// </summary>
                                        Eptsave();
                                    }
                                }
                            }
                            Sound.Play("ok");

                        }
                        else
                        {
                            Sound.Play("error");
                            return;
                        }

                    }
                }
                else
                {
                    Sound.Play("error");
                    return;
                }
                EptProcess();
                MessageBox.Show(" 出货报表处理完成 ", " 系统提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);
            });
        }
    }
}
