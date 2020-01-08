using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.Common;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ImportExcellExe
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        public class MyApplication
        {
            public static string connectionString;
            public static string connectionStringMaster;
            public static string strServerName;
            public static string strAppUserName;
        }
        public class MyDatabase
        {
            public static string strUserName;
            public static string strPassword;
        }
        public string strGodown { get; set; }
        public long strGodownSerial { get; set; }
        public string Itemname { get; set; }
        public long ItemSerial { get; set; }
        public string strBranchID { get; set; }
        public long dynamicSerial { get; set; }
        public string Process_Name { get; set; }
        public string Expenses { get; set; }
        public string product_Name { get; set; }

        public string AltUnitFirst { get; set; }
        public double AltQtyFirst { get; set; }
        public decimal BaseQtyFirst { get; set; }
        public string AltUnitSecond { get; set; }
        public double AltQtySecond { get; set; }
        public decimal BaseQtySecond { get; set; }
        public string AltUnitThird { get; set; }
        public double AltQtyThird { get; set; }
        public double BaseQtyThird { get; set; }
        public double coEfficientlngloop { get; set; }
        private void buttonSelectFile_Click(object sender, EventArgs e)
        {
            var filePath = string.Empty;
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.InitialDirectory = "c:\\";
                openFileDialog.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
                openFileDialog.FilterIndex = 2;
                openFileDialog.RestoreDirectory = true;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    //Get the path of specified file
                    filePath = openFileDialog.FileName;
                    FileTextBox.Text = filePath;

                }
            }
        }
        public string gGetServerName()
        {
            string strServerName = null;
            string strPath = Environment.CurrentDirectory;
            string FILE_NAME = strPath + @"\server.txt";
            if (System.IO.File.Exists(FILE_NAME) == true)
            {
                System.IO.StreamReader objReader = new System.IO.StreamReader(FILE_NAME);
                strServerName = (objReader.ReadLine());
                objReader.Close();
                objReader.Dispose();
            }
            else
            {
                //FileStream objFile = new FileStream(FILE_NAME, FileMode.Create, FileAccess.Write);
                //StreamWriter objWriter = new StreamWriter(objFile);
                //strServerName = Interaction.InputBox("Input a Valid Server Name", "Server Name");
                //objWriter.Write(strServerName);
                //objWriter.Close();
            }
            return strServerName;
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            ConpanyIDlistBox.DisplayMember = "Text";
            ConpanyIDlistBox.ValueMember = "Value";
            mFillCompany();
        }
        private void mFillCompany()
        {
            string CompanyName = null;
            List<string> companyList = new List<string>();
            MyApplication.strServerName = gGetServerName();
            MyDatabase.strUserName = "sa";
            MyDatabase.strPassword = "manager";
            string sqlConnectionString = ("Data Source=" + MyApplication.strServerName + ";User ID=" + MyDatabase.strUserName + ";Password=" + MyDatabase.strPassword + ";Trusted_Connection=False;");
            string queryString = string.Empty;
            ConpanyIDlistBox.Items.Clear();
            queryString = "SELECT name,RIGHT(name,4) as serial FROM master.dbo.sysdatabases where name like 'TROV%' ORDER BY serial ASC ";
            using (SqlConnection connection = new SqlConnection(sqlConnectionString))
            {
                SqlCommand command = new SqlCommand(queryString, connection);
                connection.Open();
                SqlDataReader reader = command.ExecuteReader();
                if (reader.HasRows)
                {
                    while (reader.Read())
                    {
                        companyList.Add(reader["name"].ToString());

                    }
                }
                else
                {

                }
                reader.Close();
            }


            foreach (var idString in companyList)
            {
                queryString = "SELECT COMPANY_NAME FROM " + idString + ".dbo.ACC_COMPANY";
                using (SqlConnection conn = new SqlConnection(sqlConnectionString))
                {
                    SqlCommand cmd = new SqlCommand(queryString, conn);
                    try
                    {
                        conn.Open();
                        CompanyName = (string)cmd.ExecuteScalar();
                        ConpanyIDlistBox.Items.Add(new { Text = CompanyName, Value = idString });


                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                }
            }


        }

        private void buttonImportHSCode_Click(object sender, EventArgs e)
        {
            string queryString = null;
            double customDuty = 0;
            double rDuty = 0;
            double sDuty = 0;
            double aVAT = 0;
            double VAT = 0;

            dynamic item = ConpanyIDlistBox.Items[ConpanyIDlistBox.SelectedIndex];

            var companyIDString = item.Value;

            List<string> sqlCommandList = new List<string>();

            if (companyIDString.Length < 8)
            {
                MessageBox.Show("Invalid company id");
                return;
            }
            MyApplication.strServerName = gGetServerName();
            MyDatabase.strUserName = "sa";
            MyDatabase.strPassword = "manager";
            string sqlConnectionString = ("Data Source=" + MyApplication.strServerName + ";Initial Catalog=" + companyIDString + ";User ID=" + MyDatabase.strUserName + ";Password=" + MyDatabase.strPassword + ";");

            SqlConnection SC = new SqlConnection(sqlConnectionString);
            string cmdText = @"IF EXISTS(SELECT * FROM INFORMATION_SCHEMA.TABLES 
                       WHERE TABLE_NAME = 'ACC_SERVICE_CODE') SELECT 1 ELSE SELECT 0";
            SC.Open();
            SqlCommand DateCheck = new SqlCommand(cmdText, SC);
            int x = Convert.ToInt32(DateCheck.ExecuteScalar());
            if (x == 0)
            {
                queryString = "CREATE TABLE ACC_SERVICE_CODE(";
                queryString += "CODE_SERIAL numeric(18,0) IDENTITY (1,1) NOT NULL,";
                queryString += "CODE_NAME varchar(50) CONSTRAINT PK_ACC_SERICE_CODE PRIMARY KEY,";
                queryString += "CODE_DESC varchar(250) NOT NULL,";
                queryString += "CODE_SD numeric(18,5) default 0 NOT NULL,";
                queryString += "CODE_VAT numeric(18,5) default 0 NOT NULL)";
                sqlCommandList.Add(queryString);
            }
            SC.Close();


            //Check the Content Type of the file 
            try
            {
                //Save file path 
                string path = FileTextBox.Text;
                //Save File as Temp then you can delete it if you want 
                //FileUpload1.SaveAs(path);
                //string path = @"C:\Users\Johnney\Desktop\ExcelData.xls"; 
                //For Office Excel 2010  please take a look to the followng link  http://social.msdn.microsoft.com/Forums/en-US/exceldev/thread/0f03c2de-3ee2-475f-b6a2-f4efb97de302/#ae1e6748-297d-4c6e-8f1e-8108f438e62e 
                string excelConnectionString = string.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=Excel 8.0", path);



                // Create Connection to Excel Workbook 
                using (OleDbConnection connection =
                             new OleDbConnection(excelConnectionString))
                {
                    OleDbCommand command = new OleDbCommand
                            ("Select * FROM [Sheet1$]", connection);

                    connection.Open();

                    // Create DbDataReader to Data Worksheet 
                    using (DbDataReader reader = command.ExecuteReader())
                    {
                        if (reader.HasRows)
                        {
                            while (reader.Read())
                            {
                                //string codeString = reader["ServiceCode"].ToString().Replace(" ", "");

                                string codeString = reader.GetString(1);

                                string DescString = reader.GetString(2).Replace("'", "''");


                                if (!reader.IsDBNull(reader.GetOrdinal("SD")))
                                {
                                    sDuty = Convert.ToDouble(reader["SD"].ToString());
                                }

                                if (!reader.IsDBNull(reader.GetOrdinal("VAT")))
                                {
                                    VAT = Convert.ToDouble(reader["VAT"].ToString());
                                }
                                queryString = "INSERT INTO ACC_SERVICE_CODE(CODE_NAME,CODE_DESC,CODE_SD,CODE_VAT) ";
                                queryString += "VALUES('" + codeString + "','" + DescString + "'," + sDuty + ",";
                                queryString += VAT + "";
                                queryString += ")";
                                sqlCommandList.Add(queryString);
                                customDuty = 0;
                                rDuty = 0;
                                sDuty = 0;
                                aVAT = 0;
                                VAT = 0;
                            }

                        }
                        else
                        {
                            Label1.Text = "No rows found.";
                        }

                    }
                }
            }

            catch (Exception ex)
            {
                Label1.Text = ex.Message;
            }
            SqlTransaction trans = null;
            try
            {
                SqlConnection connection1 = new SqlConnection(sqlConnectionString);

                connection1.Open();

                trans = connection1.BeginTransaction();

                foreach (var commandString in sqlCommandList)
                {
                    queryString = commandString;
                    SqlCommand command = new SqlCommand(commandString, connection1, trans);
                    command.ExecuteNonQuery();
                }

                trans.Commit();

                Label1.Text = "The data has been Imported from Excel to SQL";
                AllCrear();
                //Thread.Sleep(6000);
                //Label1.Hide();
            }
            catch (Exception ex) //error occurred
            {
                //Label1.Text = "The data has been never exported Again from Excel to SQL";
                Label1.Text = ex.ToString();
                trans.Rollback();
            }
        }

        private void AllCrear()
        {
            FileTextBox.Text = string.Empty;
        }

        private void buttonCustomer_Click(object sender, EventArgs e)
        {
            string companyName = string.Empty;
            string queryString = null;
            double customerType = 0;
            string ledgerString = string.Empty;
            double ledgerBinNo = 0;
            double ledgerTinNo = 0;
            string address1 = string.Empty;
            string address2 = string.Empty;
            string city = string.Empty;
            string contactPerson = string.Empty;
            string phoneNo = string.Empty;
            string eMail = string.Empty;
            string nationalID = string.Empty;
            string customerTypeNmae = string.Empty;
            string priceLevel = string.Empty;
            string country = string.Empty;
            string postalCode = string.Empty;
            dynamic item = ConpanyIDlistBox.Items[ConpanyIDlistBox.SelectedIndex];

            var companyIDString = item.Value;

            List<string> sqlCommandList = new List<string>();

            if (companyIDString.Length < 8)
            {
                MessageBox.Show("Invalid company id");
                return;
            }
            MyApplication.strServerName = gGetServerName();
            MyDatabase.strUserName = "sa";
            MyDatabase.strPassword = "manager";
            string sqlConnectionString = ("Data Source=" + MyApplication.strServerName + ";Initial Catalog=" + companyIDString + ";User ID=" + MyDatabase.strUserName + ";Password=" + MyDatabase.strPassword + ";");

            //Check the Content Type of the file 
            try
            {
                //Save file path 
                string path = FileTextBox.Text;
                //Save File as Temp then you can delete it if you want 
                //FileUpload1.SaveAs(path);
                //string path = @"C:\Users\Johnney\Desktop\ExcelData.xls"; 
                //For Office Excel 2010  please take a look to the followng link  http://social.msdn.microsoft.com/Forums/en-US/exceldev/thread/0f03c2de-3ee2-475f-b6a2-f4efb97de302/#ae1e6748-297d-4c6e-8f1e-8108f438e62e 
                string excelConnectionString = string.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=Excel 8.0", path);



                // Create Connection to Excel Workbook 
                using (OleDbConnection connection =
                             new OleDbConnection(excelConnectionString))
                {
                    OleDbCommand command = new OleDbCommand
                            ("Select * FROM [Sheet1$]", connection);

                    connection.Open();

                    // Create DbDataReader to Data Worksheet 
                    using (DbDataReader reader = command.ExecuteReader())
                    {
                        if (reader.HasRows)
                        {
                            while (reader.Read())
                            {
                                if (!reader.IsDBNull(reader.GetOrdinal(" Customer Name")))
                                {
                                    ledgerString = reader[" Customer Name"].ToString().Replace("'", "''");

                                }
                                if (!reader.IsDBNull(reader.GetOrdinal("BIN No")))
                                {
                                    ledgerBinNo = Convert.ToDouble(reader["BIN No"].ToString());
                                }
                                if (!reader.IsDBNull(reader.GetOrdinal("TIN No")))
                                {
                                    ledgerTinNo = Convert.ToDouble(reader["TIN No"].ToString());
                                }
                                if (!reader.IsDBNull(reader.GetOrdinal("Address1")))
                                {
                                    string address = reader["Address1"].ToString().Replace("'", "''");

                                    if (address.Length != null && address.Length >= 50)
                                    {
                                        address1 = address.Substring(0, 45);
                                        address2 = address.Substring(45);
                                    }
                                    else
                                    {
                                        address1 = reader["Address1"].ToString().Replace("'", "''");
                                    }
                                    //if (address.Length == 50)
                                    //{
                                    //   address1 = address;
                                    //}
                                }
                                if (!reader.IsDBNull(reader.GetOrdinal("Address2")))
                                {
                                    address2 = reader["Address2"].ToString();
                                }
                                if (!reader.IsDBNull(reader.GetOrdinal("City")))
                                {
                                    city = reader["City"].ToString().Replace("'", "''");
                                }
                                if (!reader.IsDBNull(reader.GetOrdinal("Postal Code")))
                                {
                                    postalCode = reader["Postal Code"].ToString();
                                }
                                if (!reader.IsDBNull(reader.GetOrdinal("Country")))
                                {
                                    country = reader["Country"].ToString();
                                }
                                if (!reader.IsDBNull(reader.GetOrdinal("Contract Person")))
                                {
                                    contactPerson = reader["Contract Person"].ToString().Replace("'", "''");
                                }
                                if (!reader.IsDBNull(reader.GetOrdinal("Phone Number")))
                                {
                                    phoneNo = reader["Phone Number"].ToString(); ;
                                }
                                if (!reader.IsDBNull(reader.GetOrdinal("E-mail")))
                                {
                                    eMail = reader["E-mail"].ToString();
                                }
                                if (!reader.IsDBNull(reader.GetOrdinal("NID No")))
                                {

                                    string nlID = reader["NID No"].ToString();

                                    if (nlID.Length != null && nlID.Length >= 17)
                                    {
                                        nationalID = nlID.Substring(0, 17);
                                    }
                                    else
                                    {
                                        nationalID = reader["NID No"].ToString();
                                    }
                                }

                                if (!reader.IsDBNull(reader.GetOrdinal("Customer Type")))
                                {
                                    customerTypeNmae = reader["Customer Type"].ToString();
                                }
                                if (customerTypeNmae == "Registered")
                                {
                                    customerType = 1;
                                }
                                else if (customerTypeNmae == "Unregistered")

                                {
                                    customerType = 3;
                                }
                                else if (customerTypeNmae == "Turn Over")
                                {
                                    customerType = 2;
                                }
                                else if (customerTypeNmae == "Foreign")
                                {
                                    customerType = 4;
                                }
                                if (!reader.IsDBNull(reader.GetOrdinal("Price Level")))
                                {
                                    priceLevel = reader["Price Level"].ToString();
                                }
                                string parentString = "Sundry Debtors";

                                string entryDate = DateTime.Now.ToString("dd/MMM/yyyy");
                                string BranchName = reader["Branch Name"].ToString();

                                gstrGetBranchID(BranchName, companyIDString);

                                queryString = "INSERT INTO ACC_LEDGER(LEDGER_BIN,LEDGER_VAT_REG_NO,LEDGER_NATIONAL_ID,LEDGER_TIN,";
                                queryString += "LEDGER_NAME,LEDGER_CASH_FLOW_TYPE,LEDGER_PARENT_GROUP,LEDGER_PRIMARY_GROUP,LEDGER_ONE_DOWN, ";
                                queryString += "LEDGER_OPENING_BALANCE,LEDGER_CLOSING_BALANCE,";
                                queryString += "LEDGER_CREDIT_LIMIT,LEDGER_CREDIT_PERIOD,LEDGER_ADDRESS1,LEDGER_ADDRESS2,LEDGER_CITY,";
                                queryString += "LEDGER_COUNTRY,SUPPLIER_TYPE,LEDGER_CONTACT,LEDGER_POSTAL,LEDGER_PHONE,LEDGER_FAX,LEDGER_EMAIL,";
                                queryString += "LEDGER_COMMENTS,LEDGER_BILL_WISE,LEDGER_STATUS,";
                                queryString += "LEDGER_LEVEL,LEDGER_GROUP,LEDGER_PRIMARY_TYPE,LEDGER_VECTOR,LEDGER_CURRENCY_SYMBOL";
                                queryString += ",LEDGER_ADD_DATE,LEDGER_PRICE_LABEL,BRANCH_ID)";
                                queryString += "VALUES('" + ledgerBinNo + "',NULL,'" + nationalID + "','" + ledgerTinNo + "','" + ledgerString + "',1,'" + parentString + "','Current Asset','" + parentString + "',";
                                queryString += "0,0,0,0,'" + address1 + "','" + address2 + "','" + city + "','" + country + "','" + customerType + "','" + contactPerson + "','" + postalCode + "','" + phoneNo + "',NULL,'" + eMail + "',";
                                queryString += "NULL,1,0,2,202,1,1,'BDT',convert(datetime,'" + entryDate + "',103) ,'" + priceLevel + "','" + strBranchID + "')";
                                sqlCommandList.Add(queryString);


                                queryString = "INSERT INTO ACC_LEDGER_TO_GROUP(GR_NAME,LEDGER_NAME) VALUES('" + parentString + "','" + ledgerString + "')";
                                sqlCommandList.Add(queryString);

                                queryString = "INSERT INTO ACC_LEDGER_TO_GROUP(GR_NAME,LEDGER_NAME) VALUES('Current Asset','" + ledgerString + "')";
                                sqlCommandList.Add(queryString);

                                queryString = "INSERT INTO ACC_LEDGER_TO_GROUP(GR_NAME,LEDGER_NAME) VALUES('Asset','" + ledgerString + "')";
                                sqlCommandList.Add(queryString);

                                //queryString = "INSERT INTO ACC_BRANCH_LEDGER_OPENING(BRANCH_LEDGER_KEY,BRANCH_ID,LEDGER_NAME,BRANCH_LEDGER_OPENING_BALANCE)";
                                //queryString += "VALUES ('" + ledgerString + "0001','0001','" + ledgerString + "',0)";
                                //sqlCommandList.Add(queryString);
                                ledgerString = string.Empty;
                                ledgerBinNo = 0;
                                ledgerTinNo = 0;
                                nationalID = null;

                            }

                        }
                        else
                        {
                            Label1.Text = "No rows found.";
                        }

                    }
                }
            }

            catch (Exception ex)
            {
                Label1.Text = ex.Message;
            }
            SqlTransaction trans = null;
            try
            {
                SqlConnection connection1 = new SqlConnection(sqlConnectionString);

                connection1.Open();

                trans = connection1.BeginTransaction();

                foreach (var commandString in sqlCommandList)
                {
                    queryString = commandString;
                    SqlCommand command = new SqlCommand(commandString, connection1, trans);
                    command.ExecuteNonQuery();
                }

                trans.Commit();

                Label1.Text = "The Customer Name has been exported succefuly from Excel to SQL";
                AllCrear();
            }
            catch (Exception ex) //error occurred
            {
                Label1.Text = "The Customer Name has not been exported Again from Excel to SQL";
                //Label1.Text = ex.ToString();
                trans.Rollback();
            }
        }
        private string gstrGetBranchID(string BranchName, dynamic companyIDString)
        {
            string queryString = string.Empty;
            //string Expenses = string.Empty;
            strBranchID = null;

            //dynamic item = ConpanyIDlistBox.Items[ConpanyIDlistBox.SelectedIndex];

            //var companyIDString = item.Value;

            List<string> sqlCommandList = new List<string>();

            MyApplication.strServerName = gGetServerName();
            MyDatabase.strUserName = "sa";
            MyDatabase.strPassword = "manager";
            string sqlConnectionString = ("Data Source=" + MyApplication.strServerName + ";Initial Catalog=" + companyIDString + ";User ID=" + MyDatabase.strUserName + ";Password=" + MyDatabase.strPassword + ";");

            ConpanyIDlistBox.Items.Clear();
            queryString = "SELECT BRANCH_ID FROM ACC_BRANCH WHERE BRANCH_NAME = '" + BranchName + "' ";
            using (SqlConnection connection = new SqlConnection(sqlConnectionString))
            {
                SqlCommand command = new SqlCommand(queryString, connection);
                connection.Open();
                SqlDataReader reader = command.ExecuteReader();
                if (reader.Read() == true)
                {
                    strBranchID = reader["BRANCH_ID"].ToString();
                }
                reader.Close();

            }
            return BranchName;
        }

        private void buttonItemName_Click(object sender, EventArgs e)
        {
            string companyName = string.Empty;
            string parentString = string.Empty;
            string ItemBaseUnit = string.Empty;
            string ItemAltUnit = string.Empty;
            string hSCode = null;
            string tarifTypeName = string.Empty;
            string queryString = null;
            string ItemString = null;
            int lngloopunit = 1;
            double ItemAltQty = 0;
            decimal ItemBaseQty = 0;
            double tarifType = 0;

            dynamic item = ConpanyIDlistBox.Items[ConpanyIDlistBox.SelectedIndex];

            var companyIDString = item.Value;

            List<string> sqlCommandList = new List<string>();

            if (companyIDString.Length < 8)
            {
                MessageBox.Show("Invalid company id");
                return;
            }
            MyApplication.strServerName = gGetServerName();
            MyDatabase.strUserName = "sa";
            MyDatabase.strPassword = "manager";
            string sqlConnectionString = ("Data Source=" + MyApplication.strServerName + ";Initial Catalog=" + companyIDString + ";User ID=" + MyDatabase.strUserName + ";Password=" + MyDatabase.strPassword + ";");

            //Check the Content Type of the file 
            try
            {
                //Save file path 

                string path = FileTextBox.Text;
                //Save File as Temp then you can delete it if you want 
                //FileUpload1.SaveAs(path);
                //string path = @"C:\Users\Johnney\Desktop\ExcelData.xls"; 
                //For Office Excel 2010  please take a look to the followng link  http://social.msdn.microsoft.com/Forums/en-US/exceldev/thread/0f03c2de-3ee2-475f-b6a2-f4efb97de302/#ae1e6748-297d-4c6e-8f1e-8108f438e62e 
                string excelConnectionString = string.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=Excel 8.0", path);



                // Create Connection to Excel Workbook 
                using (OleDbConnection connection =
                             new OleDbConnection(excelConnectionString))
                {
                    OleDbCommand command = new OleDbCommand
                            ("Select * FROM [Sheet1$]", connection);

                    connection.Open();

                    // Create DbDataReader to Data Worksheet 
                    using (DbDataReader reader = command.ExecuteReader())
                    {
                        if (reader.HasRows)
                        {
                            while (reader.Read())
                            {
                                if (!reader.IsDBNull(reader.GetOrdinal("Stock Item Name")))
                                {
                                    ItemString = reader["Stock Item Name"].ToString().Replace("'", "''");
                                    //Itemname = ItemString;
                                }
                                if (!reader.IsDBNull(reader.GetOrdinal("Under")))
                                {
                                    parentString = reader["Under"].ToString();
                                }
                                if (!reader.IsDBNull(reader.GetOrdinal("Unit")))
                                {
                                    ItemBaseUnit = reader["Unit"].ToString();
                                }
                                if (!reader.IsDBNull(reader.GetOrdinal("HS Code")))
                                {
                                    hSCode = reader["HS Code"].ToString();
                                }
                                if (!reader.IsDBNull(reader.GetOrdinal("VAT Calculation")))
                                {
                                    tarifTypeName = reader["VAT Calculation"].ToString(); ;
                                }
                                if (tarifTypeName == "By Value")
                                {
                                    tarifType = 0;
                                }
                                if (tarifTypeName == "By Quantity")
                                {
                                    tarifType = 1;
                                }
                                if (tarifTypeName == "VAT Exempted")
                                {
                                    tarifType = 2;
                                }
                                // First AltUnit

                                if (!reader.IsDBNull(reader.GetOrdinal("Alt Unit First")))
                                {
                                    AltUnitFirst = reader["Alt Unit First"].ToString();
                                }
                                if (!reader.IsDBNull(reader.GetOrdinal("Alt Qty First")))
                                {
                                    AltQtyFirst = Convert.ToDouble(reader["Alt Qty First"].ToString());
                                }
                                if (!reader.IsDBNull(reader.GetOrdinal("Base Qty First")))
                                {
                                    BaseQtyFirst = decimal.Parse(reader["Base Qty First"].ToString());
                                }

                                // Second AltUnit

                                if (!reader.IsDBNull(reader.GetOrdinal("Alt Unit Second")))
                                {
                                    AltUnitSecond = reader["Alt Unit Second"].ToString();
                                }
                                if (!reader.IsDBNull(reader.GetOrdinal("Alt Qty Second")))
                                {
                                    AltQtySecond = Convert.ToDouble(reader["Alt Qty Second"].ToString());
                                }
                                if (!reader.IsDBNull(reader.GetOrdinal("Base Qty Second")))
                                {
                                    BaseQtySecond = decimal.Parse(reader["Base Qty Second"].ToString());
                                }

                                // Third AltUnit

                                if (!reader.IsDBNull(reader.GetOrdinal("Alt Unit Third")))
                                {
                                    AltUnitThird = reader["Alt Unit Third"].ToString();
                                }
                                if (!reader.IsDBNull(reader.GetOrdinal("Alt Qty Third")))
                                {
                                    AltQtyThird = Convert.ToDouble(reader["Alt Qty Third"].ToString());
                                }
                                if (!reader.IsDBNull(reader.GetOrdinal("Base Qty Third")))
                                {
                                    BaseQtyThird = Convert.ToDouble(reader["Base Qty Third"].ToString());
                                }

                                string entryDate = DateTime.Now.ToString("dd/MMM/yyyy");

                                if (ItemString != "" && parentString != "")
                                {
                                    queryString = "INSERT INTO INV_STOCKITEM(STOCKITEM_NAME,STOCKITEM_ALIAS,STOCKITEM_DESCRIPTION,STOCKGROUP_NAME,";
                                    queryString += "STOCKITEM_PRIMARY_GROUP,STOCKCATEGORY_NAME,STOCKITEM_MANUFACTURER,STOCKITEM_BASEUNITS,";
                                    queryString += "STOCKITEM_ADDITIONALUNITS,STOCKITEM_CONVERSION,STOCKITEM_DENOMINATOR,";
                                    queryString += "STOCKITEM_OPENING_BALANCE,STOCKITEM_OPENING_RATE,STOCKITEM_OPENING_VALUE,";
                                    queryString += "STOCKITEM_MIN_QUANTITY,STOCKITEM_REORDER_LEVEL,STOCKITEM_MAINTAIN_SERIAL,HS_CODE,";
                                    queryString += "PERCENTAGE_OF_REBATE,PERCENTAGE_OF_VAT,VAT_EXEMPTED,TARIFF_TYPE,";
                                    queryString += "SD_TYPE,SD_RATE,STOCKITEM_STATUS)";
                                    queryString += "VALUES('" + ItemString + "', NULL,NULL,'" + parentString + "','" + parentString + "',NULL,NULL,'" + ItemBaseUnit + "',NULL,0,";
                                    queryString += "0,0,0,0,0,0,0,'" + hSCode + "',0,0,0,'" + tarifType + "',0,0,0)";
                                    sqlCommandList.Add(queryString);





                                    //queryString = "INSERT INTO INV_STOCKCATEGORY(STOCKCATEGORY_NAME,STOCKCATEGORY_PARENT,STOCKCATEGORY_PRIMARY,STOCKCATEGORY_OPENING_BALANCE,STOCKCATEGORY_CLOSING_BALANCE,STOCKCATEGORY_INWARDQUANTITY,STOCKCATEGORY_OUTWARDQUANTITY,STOCKCATEGORY_DEBIT_CLOSING_BAL,STOCKCATEGORY_TYPE,INSERT_DATE,EXPORT_TYPE) ";
                                    //queryString += "VALUES('" + parentString + "','" + parentString + "','" + parentString + "',0,0,0,0,0,1,convert(datetime,'" + entryDate + "',103),1)";
                                    //sqlCommandList.Add(queryString);

                                    //queryString = "INSERT INTO INV_STOCKGROUP(STOCKGROUP_NAME,STOCKGROUP_PARENT,STOCKGROUP_ONE_DOWN,STOCKGROUP_PRIMARY,STOCKGROUP_LEVEL,STOCKGROUP_SEQUENCES,STOCKGROUP_PRIMARY_TYPE,STOCKGROUP_SECONDARY_TYPE, STOCKGROUP_DEFAULT,STOCKGROUP_NAME_DEFAULT) ";
                                    //queryString += "VALUES('" + parentString + "','" + parentString + "','" + parentString + "','" + parentString + "',1,990,1,1,1,'" + parentString + "')";
                                    //sqlCommandList.Add(queryString);

                                    queryString = "INSERT INTO INV_STOCKITEM_TO_GROUP(STOCKGROUP_NAME, STOCKITEM_NAME) ";
                                    queryString += "VALUES('" + parentString + "',";
                                    queryString += "'" + ItemString + "')";
                                    sqlCommandList.Add(queryString);

                                    queryString = "INSERT INTO INV_STOCKITEM_LEVEL(STOCKITEM_NAME,STOCKGROUP_LEVEL_1) ";
                                    queryString += "VALUES('" + ItemString + "','" + parentString + "')";
                                    sqlCommandList.Add(queryString);



                                    long lngLoop = 1;
                                    multilocation(companyIDString);

                                    queryString = "INSERT INTO INV_STOCKITEM_CLOSING(STOCKITEM_NAME,GODOWNS_NAME,INSERT_DATE,EXPORT_TYPE) ";
                                    queryString += "VALUES('" + ItemString + "','" + strGodown + "',convert(datetime,'" + entryDate + "',103),1)";
                                    sqlCommandList.Add(queryString);

                                    string strRefNo = "OP" + strBranchID + ItemSerial + "-OPN" + lngLoop + strGodownSerial;

                                    queryString = "INSERT INTO INV_MASTER(INV_REF_NO,INV_DATE,INWORD_QUANTITY,INV_OPENING_FLAG,BRANCH_ID) ";
                                    queryString += "VALUES('" + strRefNo + "', convert(datetime,'" + entryDate + "',103) ,";
                                    queryString += " 0,1,";
                                    queryString += "'" + strBranchID + "'";
                                    queryString += ")";
                                    sqlCommandList.Add(queryString);

                                    queryString = "INSERT INTO INV_TRAN(INV_TRAN_KEY,INV_TRAN_POSITION,BRANCH_ID,INV_REF_NO,INV_DATE,STOCKITEM_NAME,";
                                    queryString += "INV_TRAN_QUANTITY,INV_UOM,INV_PER,INV_TRAN_RATE,INV_TRAN_AMOUNT,GODOWNS_NAME,";
                                    queryString += "INV_LOG_NO,INV_VOUCHER_TYPE,INV_OPENING_FLAG) ";
                                    queryString += "VALUES('" + strRefNo + "',";
                                    queryString += "" + lngLoop + ",'" + strBranchID + "',";
                                    queryString += "'" + strRefNo + "',convert(datetime,'" + entryDate + "',103),";
                                    queryString += "'" + ItemString + "',0,'" + ItemBaseUnit + "','" + ItemBaseUnit + "',";
                                    queryString += " 0,0,";
                                    queryString += "'" + strGodown + "',";
                                    //If uctxtOpeningBatch.Text <> vbNullString Then
                                    //    If Trim$(uctxtOpeningBatch.Text) <> gcEND_OF_LIST Then
                                    queryString += "NULL,";
                                    //    End If
                                    //End If
                                    queryString += "0,1)";
                                    sqlCommandList.Add(queryString);

                                }

                                //First Alt Unit
                                queryString = "INSERT INTO INV_STOCKITEM_UOM(STOCKITEM_NAME,STOCKITEM_UOM_POSITION,STOCKITEM_UNIT,STOCKITEM_CONVERSION,STOCKITEM_DENOMINATOR) ";
                                queryString += "VALUES('" + ItemString + "',1,'" + ItemBaseUnit + "',1,1)";
                                sqlCommandList.Add(queryString);

                                if (ItemString != "" && AltUnitFirst != "" && AltQtyFirst != 0)
                                {
                                    lngloopunit = lngloopunit + 1;
                                    queryString = "INSERT INTO INV_STOCKITEM_UOM(STOCKITEM_NAME,STOCKITEM_UOM_POSITION,STOCKITEM_UNIT,STOCKITEM_CONVERSION,STOCKITEM_DENOMINATOR) ";
                                    queryString += "VALUES('" + ItemString + "','" + lngloopunit + "','" + AltUnitFirst + "','" + AltQtyFirst + "','" + BaseQtyFirst + "')";
                                    sqlCommandList.Add(queryString);
                                }

                                //Second Alt Unit

                                if (ItemString != "" && AltUnitSecond != "" && AltQtySecond != 0)
                                {
                                    lngloopunit = lngloopunit + 1;
                                    queryString = "INSERT INTO INV_STOCKITEM_UOM(STOCKITEM_NAME,STOCKITEM_UOM_POSITION,STOCKITEM_UNIT,STOCKITEM_CONVERSION,STOCKITEM_DENOMINATOR) ";
                                    queryString += "VALUES('" + ItemString + "','" + lngloopunit + "','" + AltUnitSecond + "','" + AltQtySecond + "','" + BaseQtySecond + "')";
                                    sqlCommandList.Add(queryString);
                                }

                                //Third Alt Unit

                                if (ItemString != "" && AltUnitThird != "" && AltQtyThird != 0)
                                {
                                    lngloopunit = lngloopunit + 1;
                                    queryString = "INSERT INTO INV_STOCKITEM_UOM(STOCKITEM_NAME,STOCKITEM_UOM_POSITION,STOCKITEM_UNIT,STOCKITEM_CONVERSION,STOCKITEM_DENOMINATOR) ";
                                    queryString += "VALUES('" + ItemString + "','" + lngloopunit + "','" + AltUnitThird + "','" + AltQtyThird + "','" + Convert.ToDecimal(BaseQtyThird) + "')";
                                    sqlCommandList.Add(queryString);
                                }


                                ItemString = string.Empty;
                                parentString = string.Empty;
                                ItemBaseUnit = string.Empty;
                                hSCode = null;
                                tarifTypeName = null;
                                //ItemAltUnit = null;
                                //ItemAltQty = 0;
                                //ItemBaseQty = 0;
                                lngloopunit = 1;
                            }

                        }
                        else
                        {
                            Label1.Text = "No rows found.";
                        }

                    }
                }
            }

            catch (Exception ex)
            {
                Label1.Text = ex.Message;
            }
            SqlTransaction trans = null;
            try
            {
                SqlConnection connection1 = new SqlConnection(sqlConnectionString);

                connection1.Open();

                trans = connection1.BeginTransaction();

                foreach (var commandString in sqlCommandList)
                {
                    queryString = commandString;
                    SqlCommand command = new SqlCommand(commandString, connection1, trans);
                    command.ExecuteNonQuery();
                }

                trans.Commit();
                Label1.Text = "The Item Name has been exported succefuly from Excel to SQL";
                AllCrear();
            }
            catch (Exception ex) //error occurred
            {
                Label1.Text = "The Item Name has not been exported Again from Excel to SQL";
                //Label1.Text = ex.ToString();
                trans.Rollback();
            }
        }

        private void multilocation(dynamic companyIDString)
        {
            string queryString = string.Empty;
            //string Expenses = string.Empty;


            //dynamic item = ConpanyIDlistBox.Items[ConpanyIDlistBox.SelectedIndex];

            //var companyIDString = item.Value;

            List<string> sqlCommandList = new List<string>();

            MyApplication.strServerName = gGetServerName();
            MyDatabase.strUserName = "sa";
            MyDatabase.strPassword = "manager";
            string sqlConnectionString = ("Data Source=" + MyApplication.strServerName + ";Initial Catalog=" + companyIDString + ";User ID=" + MyDatabase.strUserName + ";Password=" + MyDatabase.strPassword + ";");

            ConpanyIDlistBox.Items.Clear();
            queryString = "Select * From INV_GODOWNS WHERE GODOWNS_DEFAULT = '1' ";
            using (SqlConnection connection = new SqlConnection(sqlConnectionString))
            {
                SqlCommand command = new SqlCommand(queryString, connection);
                connection.Open();
                SqlDataReader reader = command.ExecuteReader();
                if (reader.Read() == true)
                {
                    strGodown = reader["GODOWNS_NAME"].ToString();
                    strBranchID = reader["BRANCH_ID"].ToString();
                }
                reader.Close();

                queryString = "SELECT GODOWNS_SERIAL FROM INV_GODOWNS WHERE GODOWNS_NAME = '" + strGodown + "' ";
                command = new SqlCommand(queryString, connection);
                reader = command.ExecuteReader();
                if (reader.Read() == true)
                {
                    strGodownSerial = Convert.ToInt32(reader["GODOWNS_SERIAL"].ToString());
                }
                reader.Close();
                //queryString = "SELECT STOCKITEM_SERIAL FROM INV_STOCKITEM  WHERE STOCKITEM_NAME = '" + Itemname + "' ";

                queryString = "SELECT STOCKITEM_SERIAL FROM INV_STOCKITEM  ORDER BY STOCKITEM_SERIAL DESC";

                command = new SqlCommand(queryString, connection);
                reader = command.ExecuteReader();
                if (reader.Read() == true)
                {
                    //dynamicSerial = 1;
                    if (ItemSerial == 0)
                    {
                        ItemSerial = Convert.ToInt32(reader["STOCKITEM_SERIAL"].ToString());
                        //ItemSerial = dynamicSerial;

                        //dynamicSerial += 1;
                    }
                }
                //else
                //{
                //    ItemSerial = dynamicSerial + 1;
                //}
                reader.Close();
                //If gobjEnhance.SearchRecord(strSQL, rsGet, adLockReadOnly, gcnMain) Then
                //    strItemSerial = rsGet.Fields("STOCKITEM_SERIAL").Value
                //End If
                ItemSerial += 1;
            }
        }

        private void buttonSupplier_Click(object sender, EventArgs e)
        {

        }
    }
}
