//ecode length needs to be specified to every single eproducttype that needs excat digit code
//specify in Eproduct() and CheckDuplicatedCode()

using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Messaging;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using MySql.Data.MySqlClient;
using OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime;
using System.Globalization;
using System.Windows.Forms;
using System.Threading;
using System.Text;
using System.Linq.Expressions;
using System.Configuration;
using System.Web.Services.Protocols;

namespace Excel
{
    public partial class WebForm1 : System.Web.UI.Page
    {
        string strcon = ConfigurationManager.ConnectionStrings["mySql"].ConnectionString;
        protected void Page_Load(object sender, EventArgs e)
        {
            Button1.OnClientClick = "return confirm('Are you sure to import to database?')";
            if (!IsPostBack)
            {
                DataTable startuptable = new DataTable();
                MySqlConnection startupconnection = new MySqlConnection(strcon);
                startupconnection.Open();
                string startupquery = "SELECT ISKU, Title FROM eproducttype WHERE STATUS = 'A'";
                MySqlCommand startupcommand = new MySqlCommand(startupquery, startupconnection);
                MySqlDataAdapter startupadapter = new MySqlDataAdapter(startupcommand);

                startupadapter.Fill(startuptable);

                DropDownList1.DataSource = startuptable;
                DropDownList1.DataTextField = "Title";
                DropDownList1.DataValueField = "ISKU";
                DropDownList1.DataBind();

                startupconnection.Close();

                TextBox2.Text = DropDownList1.SelectedValue.ToString();
            }

            TextBox2.Text = DropDownList1.SelectedValue.ToString();
        }

        DataTable tbl = new DataTable();
        DataTable dt = new DataTable();
        string path = "";

        protected void Browse_Click(object sender, EventArgs e)
        {
            var t = new Thread((ThreadStart)(() =>
            {
                OpenFileDialog openFileDialog1 = new OpenFileDialog
                {
                    //InitialDirectory = @"C:\Desktop",
                    Title = "Browse Excel Files",

                    CheckFileExists = true,
                    CheckPathExists = true,

                    DefaultExt = "xlsx",
                    Filter = "Excel files |*.xls;*.xlsx;*.xlsm",
                    FilterIndex = 2,
                    RestoreDirectory = true,

                };

                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    path = openFileDialog1.FileName;
                    TextBox1.Text = path.ToString();
                }

            }));

            t.SetApartmentState(ApartmentState.STA);
            t.Start();
            t.Join();
        }
        protected void Import_Click(object sender, EventArgs e)
        {
            try
            {
                CheckDuplicatedCode();
                if (CheckDuplicatedCode())
                {
                    MySqlConnection con = new MySqlConnection(strcon);
                    con.Open();
                    MySqlTransaction tr = null;
                    try
                    {
                        path = TextBox1.Text.ToString();
                        tr = con.BeginTransaction();

                        Eproduct(path, true);

                        ETransaction(path, true);

                        tr.Commit();
                        Response.Write("<script>alert('Import Completed')</script>");

                    }

                    catch (MySqlException)
                    {
                        tr.Rollback();
                        Response.Write("<script>alert('MySQL.Exception.Connection Failed')</script>");
                    }

                    catch
                    {
                        tr.Rollback();
                        Response.Write("<script>alert('Connection failed')</script>");
                    }

                    con.Close();
                }

                else
                {
                    Response.Write("<script>alert('Duplicate Ecode exists')</script>");
                }
            }

            catch
            {
                Response.Write("<script>alert('Invalid Path')</script>");
            }
        }

        public DataTable Eproduct(string path, bool hasHeader = true)
        {
            using (var pck = new OfficeOpenXml.ExcelPackage())
            {
                using (var stream = File.OpenRead(path))
                {
                    pck.Load(stream);
                }

                var ws = pck.Workbook.Worksheets.First();

                var startRow = hasHeader ? 2 : 1;

                tbl.Columns.Add("ID", typeof(int));
                tbl.Columns.Add("Ecode", typeof(string));
                tbl.Columns.Add("EcodeSecondary", typeof(string));
                tbl.Columns.Add("ProdType", typeof(int));
                tbl.Columns.Add("ISKU", typeof(string));
                tbl.Columns.Add("Title", typeof(string));
                tbl.Columns.Add("CurrencyId", typeof(int));
                tbl.Columns.Add("Credit", typeof(double));
                tbl.Columns.Add("StartDate", typeof(DateTime));
                tbl.Columns.Add("EndDate", typeof(DateTime));
                tbl.Columns.Add("IsActive", typeof(int));
                tbl.Columns.Add("IsSend", typeof(int));
                tbl.Columns.Add("OrderNo", typeof(string));
                tbl.Columns.Add("POId", typeof(string));

                DataRow myNewRow;
                for (int i = startRow; i <= ws.Dimension.End.Row; i++)
                {
                    myNewRow = tbl.NewRow();
                    string ecode = ws.Cells[i, 2].Value.ToString();

                    myNewRow["EcodeSecondary"] = ws.Cells[i, 1].Value;
                    myNewRow["ISKU"] = TextBox2.Text;
                    myNewRow["Credit"] = ws.Cells[i, 4].Value;
                    myNewRow["StartDate"] = DateTime.Now;
                    myNewRow["EndDate"] = ws.Cells[i, 3].Value.ToString();

                    myNewRow["CurrencyId"] = 1;
                    myNewRow["IsActive"] = 1;
                    myNewRow["IsSend"] = 0;

                    string mySelectQuery = $"SELECT Id, Title FROM eproducttype WHERE ISKU = '{TextBox2.Text}'";
                    MySqlConnection myConnection = new MySqlConnection(strcon);
                    myConnection.Open();
                    MySqlCommand myCommand = new MySqlCommand(mySelectQuery, myConnection);

                    MySqlDataReader myReader;
                    myReader = myCommand.ExecuteReader();

                    while (myReader.Read())
                    {
                        int prodtype = myReader.GetInt32(0);
                        myNewRow["ProdType"] = prodtype;
                        string title = myReader.GetString(1);
                        string newtitle;
                        if (title.Contains("'"))
                        {
                            newtitle = title.Replace("'", "''");
                            myNewRow["Title"] = newtitle;
                        }

                        else
                        {
                            myNewRow["Title"] = myReader.GetString(1);
                        }

                        if (prodtype == 8)
                        {
                            if (ecode.Length != 10)
                            {
                                string newecode = ecode.PadLeft(10, '0');
                                myNewRow["Ecode"] = newecode;
                            }

                            else
                            {
                                myNewRow["Ecode"] = ecode;
                            }

                        }

                        else
                        {
                            myNewRow["Ecode"] = ecode;
                        }
                    }
                    myReader.Close();

                    myConnection.Close();

                    tbl.Rows.Add(myNewRow);
                }
                MySqlConnection con = new MySqlConnection(strcon);
                con.Open();
                MySqlCommand cmd;

                foreach (DataRow row in tbl.Rows)
                {
                    string startdate = row["StartDate"].ToString();
                    DateTime startdatedatetime = new DateTime(); //12/8/2020 11:44:16 AM //yyyy-MM-dd
                    startdatedatetime = DateTime.ParseExact(startdate, "d/M/yyyy h:mm:ss tt", null);

                    string enddate = row["EndDate"].ToString();
                    DateTime enddatedatetime = new DateTime(); //8/6/2020 11:44:16 AM //yyyy-MM-dd
                    //enddatedatetime = DateTime.ParseExact(enddate, "yyyy-MM-dd h:mm:ss tt", null);
                    enddatedatetime = DateTime.Parse(enddate);

                    string query = "INSERT INTO eproduct (Ecode, EcodeSecondary, ProdType, ISKU, Title, Credit, StartDate, EndDate, CurrencyId, IsActive, IsSend) " +
                        "VALUES ('" + row["Ecode"] + "','" + row["EcodeSecondary"] + "','" + row["ProdType"] + "','" + row["ISKU"] + "','" + @row["Title"] + "','" + row["Credit"] + "','" + startdatedatetime.ToString("yyyy-MM-dd HH:mm:ss") + "','" + enddatedatetime.ToString("yyyy-MM-dd HH:mm:ss") + "','" + row["CurrencyId"] + "','" + row["IsActive"] + "','" + row["IsSend"] + "')";
                    cmd = new MySqlCommand(query, con);
                    cmd.ExecuteNonQuery();
                }
                con.Close();
                return tbl;
            }
        }

        public DataTable ETransaction(string path, bool hasHeader = true)
        {
            using (var pck = new OfficeOpenXml.ExcelPackage())
            {

                using (var stream = File.OpenRead(path))
                {
                    pck.Load(stream);
                }

                var ws = pck.Workbook.Worksheets.First();

                var startRow = hasHeader ? 2 : 1;

                dt.Columns.Add("ID", typeof(int));
                dt.Columns.Add("OrderNo", typeof(string));
                dt.Columns.Add("RefNo", typeof(string));
                dt.Columns.Add("PONo", typeof(string));
                dt.Columns.Add("ISKU", typeof(string));
                dt.Columns.Add("ProdType", typeof(int));
                dt.Columns.Add("CustId", typeof(string));
                dt.Columns.Add("Email", typeof(string));
                dt.Columns.Add("SellerId", typeof(string));
                dt.Columns.Add("SkinId", typeof(string));
                dt.Columns.Add("Status", typeof(int));
                dt.Columns.Add("IsSend", typeof(int));
                dt.Columns.Add("OrderDate", typeof(DateTime));
                dt.Columns.Add("EmailSentDate", typeof(DateTime));
                dt.Columns.Add("CreatedDate", typeof(DateTime));
                dt.Columns.Add("LastUpdatedDate", typeof(DateTime));
                dt.Columns.Add("Remark", typeof(string));
                dt.Columns.Add("TransactionType", typeof(string));
                dt.Columns.Add("Qty", typeof(int));
                dt.Columns.Add("UnitCredit", typeof(double));
                dt.Columns.Add("TransactionCredit", typeof(double));
                dt.Columns.Add("BalanceCredit", typeof(double));
                dt.Columns.Add("BalQty", typeof(int));

                DataRow myNewRow;
                int qty = 0;
                for (int i = startRow; i <= ws.Dimension.End.Row; i++)
                {
                    qty++;
                }
                double unitcredit = Convert.ToInt32(ws.Cells[2, 4].Value);
                double transactioncredit = Convert.ToDouble(qty) * unitcredit;
                string remark = "IN";
                string transactiontype = "Create";

                myNewRow = dt.NewRow();
                myNewRow["ISKU"] = TextBox2.Text;
                myNewRow["Remark"] = remark;
                myNewRow["TransactionType"] = transactiontype;
                myNewRow["Qty"] = qty;
                myNewRow["UnitCredit"] = unitcredit;
                myNewRow["TransactionCredit"] = transactioncredit;

                string mySelectQuery = $"SELECT Id FROM eproducttype WHERE ISKU = '{TextBox2.Text}'";
                MySqlConnection myConnection = new MySqlConnection(strcon);
                myConnection.Open();
                MySqlCommand myCommand = new MySqlCommand(mySelectQuery, myConnection);

                MySqlDataReader myReader;
                myReader = myCommand.ExecuteReader();

                while (myReader.Read())
                {
                    myNewRow["ProdType"] = myReader.GetInt32(0);
                }
                myReader.Close();
                myConnection.Close();

                string mySqlquery = $"SELECT BalanceCredit, BalQty FROM etransaction WHERE ISKU = '{TextBox2.Text}' ORDER BY Id DESC LIMIT 1";
                MySqlConnection mySqlConnection = new MySqlConnection(strcon);
                mySqlConnection.Open();
                MySqlCommand mySqlCommand = new MySqlCommand(mySqlquery, mySqlConnection);

                MySqlDataReader mySqlDataReader;
                mySqlDataReader = mySqlCommand.ExecuteReader();

                double balancecredit = 0;
                int balqty = 0;
                while (mySqlDataReader.Read())
                {
                    balancecredit = mySqlDataReader.GetDouble(0);
                    balqty = mySqlDataReader.GetInt32(1);
                }
                mySqlDataReader.Close();

                balancecredit += transactioncredit;
                balqty += qty;

                myNewRow["BalanceCredit"] = balancecredit;
                myNewRow["BalQty"] = balqty;

                mySqlConnection.Close();

                dt.Rows.Add(myNewRow);
                MySqlConnection con = new MySqlConnection(strcon);
                con.Open();
                MySqlCommand cmd;

                foreach (DataRow row in dt.Rows)
                {
                    string query = "INSERT INTO etransaction (ISKU, ProdType, Remark, TransactionType, Qty, UnitCredit, TransactionCredit, BalanceCredit, BalQty) " +
                        "VALUES ('" + row["ISKU"] + "','" + row["ProdType"] + "','" + row["Remark"] + "','" + row["TransactionType"] + "','" + row["Qty"] + "'" +
                        ",'" + row["UnitCredit"] + "','" + row["TransactionCredit"] + "','" + row["BalanceCredit"] + "','" + row["BalQty"] + "')";
                    cmd = new MySqlCommand(query, con);
                    cmd.ExecuteNonQuery();
                }
                con.Close();
                return dt;
            }
        }

        protected void DropDownList1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        protected void Check_Click(object sender, EventArgs e)
        {
            try
            {
                path = TextBox1.Text.ToString();
                using (var pck = new OfficeOpenXml.ExcelPackage())
                {

                    using (var stream = File.OpenRead(path))
                    {
                        pck.Load(stream);
                    }

                    var ws = pck.Workbook.Worksheets.First();

                    var startRow = 2;
                    DataTable dtt = new DataTable();

                    dtt.Columns.Add("ISKU", typeof(string));
                    dtt.Columns.Add("ProdType", typeof(int));
                    dtt.Columns.Add("Remark", typeof(string));
                    dtt.Columns.Add("TransactionType", typeof(string));
                    dtt.Columns.Add("Qty", typeof(int));
                    dtt.Columns.Add("UnitCredit", typeof(double));
                    dtt.Columns.Add("TransactionCredit", typeof(double));

                    DataRow myNewRow;
                    int qty = 0;
                    for (int i = startRow; i <= ws.Dimension.End.Row; i++)
                    {
                        qty++;
                    }
                    double unitcredit = Convert.ToInt32(ws.Cells[2, 4].Value);
                    double transactioncredit = Convert.ToDouble(qty) * unitcredit;
                    string remark = "IN";
                    string transactiontype = "Create";

                    myNewRow = dtt.NewRow();
                    myNewRow["ISKU"] = TextBox2.Text;
                    myNewRow["Remark"] = remark;
                    myNewRow["TransactionType"] = transactiontype;
                    myNewRow["Qty"] = qty;
                    myNewRow["UnitCredit"] = unitcredit;
                    myNewRow["TransactionCredit"] = transactioncredit;

                    string mySelectQuery = $"SELECT Id FROM eproducttype WHERE ISKU = '{TextBox2.Text}'";
                    MySqlConnection myConnection = new MySqlConnection(strcon);
                    myConnection.Open();
                    MySqlCommand myCommand = new MySqlCommand(mySelectQuery, myConnection);

                    MySqlDataReader myReader;
                    myReader = myCommand.ExecuteReader();

                    while (myReader.Read())
                    {
                        myNewRow["ProdType"] = myReader.GetInt32(0);
                    }
                    myReader.Close();

                    dtt.Rows.Add(myNewRow);

                    myConnection.Close();

                    GridView1.DataSource = dtt;
                    GridView1.DataBind();
                }
            }

            catch
            {
                Response.Write("<script>alert('Invalid Path')</script>");
            }

        }

        protected bool CheckDuplicatedCode()
        {

            using (var pck = new OfficeOpenXml.ExcelPackage())
            {
                path = TextBox1.Text.ToString();
                using (var stream = File.OpenRead(path))
                {
                    pck.Load(stream);
                }

                var ws = pck.Workbook.Worksheets.First();

                var startRow = 2;
                string ecode = "";
                int i = startRow;

                DataTable dctbl = new DataTable();
                dctbl.Columns.Add("Duplicate Ecode");
                DataRow dcrow;

                bool duplicate = false;
                while (i <= ws.Dimension.End.Row)
                {
                    ecode = ws.Cells[i, 2].Value.ToString();

                    DataTable dataTable = new DataTable();

                    MySqlConnection mySqlConnection = new MySqlConnection(strcon);
                    mySqlConnection.Open();

                    string mySqlQuery = $"SELECT Ecode FROM eproduct WHERE ISKU = '{TextBox2.Text}'";
                    MySqlCommand mySqlCommand = new MySqlCommand(mySqlQuery, mySqlConnection);
                    MySqlDataReader mySqlDataReader;
                    mySqlDataReader = mySqlCommand.ExecuteReader();

                    if (TextBox2.Text == "418224-429275") //Touch 'n Go eWallet Reload PIN RM50
                    {
                        if (ecode.Length != 10)
                        {
                            ecode = ecode.PadLeft(10, '0');
                        }
                    }

                    while (mySqlDataReader.Read())
                    {
                        if (ecode == mySqlDataReader.GetString(0))
                        {
                            duplicate = true;
                            dcrow = dctbl.NewRow();
                            dcrow["Duplicate Ecode"] = ecode;
                            dctbl.Rows.Add(dcrow);
                        }
                    }
                    mySqlConnection.Close();
                    i++;
                }

                if (duplicate)
                {
                    GridView1.DataSource = dctbl;
                    GridView1.DataBind();
                    return false;
                }

                else
                {
                    GridView1.DataSource = null;
                    GridView1.DataBind();
                    return true;
                }
            }
        }

        protected void Upload_Click(object sender, EventArgs e)
        {
            string path = TextBox1.Text;
            string destination = Server.MapPath(@"~/Data/");

            string destinationfilename = Path.Combine(destination, Path.GetFileName(path));
            if (!File.Exists(destinationfilename))
            {
                File.Copy(path, Path.Combine(destination, Path.GetFileName(path)));
                Response.Write("<script>alert('File uploaded')</script>");
            }

            else
            {
                Response.Write("<script>alert('File exists')</script>");
            }
        }
    }
}