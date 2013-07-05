using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Text.RegularExpressions;
using System.Data.Sql;


namespace dCube
{
    public partial class FrmConnection : Form
    {
        /// <summary>
        ///     Access OLEDB Connection String Driver
        /// Provider=Microsoft.Jet.OLEDB.4.0;Data Source=\directory\demo.mdb;User Id=admin;Password=;
        ///     Excel OLEDB Connection String
        /// Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\MyExcel.xls;Extended Properties='"Excel 8.0;HDR=Yes;IMEX=1"'
        ///     FoxPro OLEDB Connection String
        /// Provider=vfpoledb.1;Data Source=c:\directory\demo.dbc;Collating Sequence=machine
        ///     MySQL OLEDB Connection String
        /// Provider=MySQLProv;Data Source=mydemodb;User Id=myusername;Password=mypasswd;
        ///     Oracle OLEDB Connection String
        /// Provider=msdaora;Data Source=mydemodb;User Id=myusername;Password=mypasswd;
        ///     SQL Server OLEDB Connection String - Database Login
        /// Provider=sqloledb;Data Source=myservername;Initial Catalog=mydemodb;User Id=myusername;Password=mypasswd;        
        /// </summary>
        public string THEME = "";
        string connection = "";
        string type = "";
        clsConnectSQL conn = new clsConnectSQL();
        public string Type
        {
            get { return type; }
            set { type = value; }
        }

        public string Connection
        {
            get { return connection; }
            set { connection = value; }
        }
        public FrmConnection()
        {
            InitializeComponent();
        }
        private void FrmConnection_Load(object sender, EventArgs e)
        {
            //if (THEME != "")
            //ThemeResolutionService.ApplyThemeToControlTree(this, THEME);


        }
        public void SetConnect(string connectstring)
        {
            if (type == "QD")
            {
                txtGeneralTimeOut.Value = 0;
                txtGeneralTimeOut.Enabled = false;
            }
            string[] arr = connectstring.Split(';');
            for (int i = 0; i < arr.Length; i++)
            {
                string[] arrP = arr[i].Split('=');
                if (arrP.Length == 2)
                {
                    if (arrP[0] == "Server")
                        Server.Text = arrP[1];
                    else if (arrP[0] == "Database")
                        Database.Text = arrP[1];
                    else if (arrP[0] == "User Id")
                        User.Text = arrP[1];
                    else if (arrP[0] == "Password")
                        Pass.Text = arrP[1];
                    else if (arrP[0] == "General Timeout" && type != "QD")
                        txtGeneralTimeOut.Text = arrP[1];
                }
            }
        }
        private void btnOKQD_Click(object sender, EventArgs e)
        {
            string timeout = "";
            if (txtGeneralTimeOut.Value != 0)
                timeout = ";General Timeout=" + txtGeneralTimeOut.Value;
            string connect = "Persist Security Info=True;Database={1};Server={0};User Id={2};Password={3}{4}";
            Connection = string.Format(connect, Server.Text, Database.Text, User.Text, Pass.Text, timeout);
            if (type != "QD")
                Connection = "Provider=SQLOLEDB.1;" + Connection;
            DialogResult = DialogResult.OK;

            Close();
        }



        public class clsConnectSQL
        {
            static String _connectString = "Data Source=[SERVER];Initial Catalog=[DATABASE]; uid=[USERNAME];pwd=[PASSWORD];Integrated Security=False";
            SqlConnection conn;

            public clsConnectSQL()
            {
                //    _connectString = "Data Source=[SERVER];Initial Catalog=[DATABASE]; uid=[USERNAME];pwd=[PASSWORD];Integrated Security=False";
                //      conn = new SqlConnection(_connectString);
            }
            public String ConnectString
            {
                get { return _connectString; }
                set
                {
                    _connectString = value;
                    conn = new SqlConnection(_connectString);
                }
            }
            public bool TestConnect()
            {
                try
                {
                    conn.Open();
                    conn.Close();
                    return true;
                }
                catch
                {
                    return false;
                }
            }
            public DataSet GetField(String tableName)
            {
                try
                {
                    //string queryField = "SELECT *,COLUMNPROPERTY(OBJECT_ID(A.TABLE_NAME),A.COLUMN_NAME,'IsIdentity') as IS_IDENTITY " +
                    //    " FROM ( SELECT D.TABLE_NAME,D.ORDINAL_POSITION, D.COLUMN_NAME,DATA_TYPE,CHARACTER_MAXIMUM_LENGTH,NUMERIC_PRECISION, NUMERIC_SCALE,MAX(CONSTRAINT_NAME) AS CONSTRAINT_NAME,IS_NULLABLE FROM INFORMATION_SCHEMA.COLUMNS D  LEFT JOIN INFORMATION_SCHEMA.KEY_COLUMN_USAGE C ON (C.COLUMN_NAME=D.COLUMN_NAME and C.TABLE_NAME=D.TABLE_NAME)  WHERE D.TABLE_NAME='" + tableName + "'" +
                    //    " GROUP BY D.TABLE_NAME,D.ORDINAL_POSITION, D.COLUMN_NAME,DATA_TYPE,CHARACTER_MAXIMUM_LENGTH,NUMERIC_PRECISION, NUMERIC_SCALE, IS_NULLABLE ) A " +
                    //    " INNER JOIN INFORMATION_SCHEMA.TABLES t ON (t.TABLE_NAME = A.TABLE_NAME AND t.TABLE_TYPE = 'BASE TABLE')";
                    string queryField = "SELECT     A.TABLE_NAME, A.ORDINAL_POSITION, A.COLUMN_NAME, A.DATA_TYPE, A.CHARACTER_MAXIMUM_LENGTH, A.NUMERIC_PRECISION, A.NUMERIC_SCALE, " +
                                                              "ISNULL(A.CONSTRAINT_NAME,'') as CONSTRAINT_NAME, A.IS_NULLABLE, t.TABLE_CATALOG, t.TABLE_SCHEMA, t.TABLE_NAME AS Expr1, t.TABLE_TYPE, " +
                                                              "ISNULL(COLUMNPROPERTY(OBJECT_ID(A.TABLE_NAME), A.COLUMN_NAME, 'IsIdentity'),'0') AS IS_IDENTITY " +
                                        "FROM         (SELECT     D.TABLE_NAME, D.ORDINAL_POSITION, D.COLUMN_NAME, D.DATA_TYPE, D.CHARACTER_MAXIMUM_LENGTH, D.NUMERIC_PRECISION, " +
                                                                                      "D.NUMERIC_SCALE, MAX(C.CONSTRAINT_NAME) AS CONSTRAINT_NAME, D.IS_NULLABLE " +
                                                               "FROM          INFORMATION_SCHEMA.COLUMNS AS D LEFT OUTER JOIN " +
                                                                                      "INFORMATION_SCHEMA.KEY_COLUMN_USAGE AS C ON C.COLUMN_NAME = D.COLUMN_NAME AND C.TABLE_NAME = D.TABLE_NAME " +
                                                               "WHERE      (D.TABLE_NAME = '" + tableName + "') " +
                                                               "GROUP BY D.TABLE_NAME, D.ORDINAL_POSITION, D.COLUMN_NAME, D.DATA_TYPE, D.CHARACTER_MAXIMUM_LENGTH, D.NUMERIC_PRECISION, " +
                                                                                      "D.NUMERIC_SCALE, D.IS_NULLABLE) AS A INNER JOIN " +
                                                              "INFORMATION_SCHEMA.TABLES AS t ON t.TABLE_NAME = A.TABLE_NAME AND t.TABLE_TYPE = 'BASE TABLE'";
                    SqlDataAdapter adap = new SqlDataAdapter(queryField, conn);
                    DataSet dset = new DataSet();
                    adap.Fill(dset);
                    return dset;
                }
                catch
                {
                    return null;
                }
            }
            public DataTable FilterFields(DataTable dt)
            {
                DataColumn parent = new DataColumn("IS_PARENT", typeof(String));
                dt.Columns.Add(parent);
                if (dt.Rows.Count > 0)
                {
                    DataRow[] rows = dt.Select("CONSTRAINT_NAME like 'PK%'");
                    foreach (DataRow row in rows)
                    {
                        row["IS_PARENT"] = "1";
                    }
                    if (rows.Length == 0)
                    {
                        DataRow[] tmpRows = dt.Select("CONSTRAINT_NAME like 'FK%'");
                        if (tmpRows.Length == 0)
                        {
                            dt.Rows[0]["CONSTRAINT_NAME"] = "PK_aaaaaa";
                            dt.Rows[0]["IS_PARENT"] = "1";
                        }
                        else
                        {
                            foreach (DataRow row in dt.Rows)
                            {
                                if (Regex.IsMatch(row["CONSTRAINT_NAME"].ToString(), @"^FK"))
                                {
                                    row["CONSTRAINT_NAME"] = "PK_aaaaaa";
                                    row["IS_PARENT"] = "1";
                                }
                            }
                        }

                    }
                }
                return dt;
            }
            public Boolean HasIndentity(String tablename)
            {
                try
                {
                    string queryField = String.Format("SELECT  OBJECTPROPERTY(OBJECT_ID('{0}'),'TableHasIdentity')", tablename);
                    SqlDataAdapter adap = new SqlDataAdapter(queryField, conn);
                    DataSet dset = new DataSet();
                    adap.Fill(dset);
                    if (Convert.ToInt32(dset.Tables[0].Rows[0][0]) == 1)
                        return true;
                    else
                        return false;
                }
                catch
                {
                    return false;
                }
            }
            public DataTable GetServers()
            {
                DataTable dt = SqlDataSourceEnumerator.Instance.GetDataSources();
                return dt;
            }
            public DataTable GetDataBases(String Server, String Username, String Pass)
            {
                try
                {
                    //Server=.;Database=SiteCamera;uid=sa;pwd=qawsed;Connection Lifetime=100;Connect Timeout=500
                    string connectString = String.Format("Server={0}; User Id={1};pwd={2}; Connection Lifetime=100;Connect Timeout=500", Server, Username, Pass);
                    conn = new SqlConnection(connectString);
                    conn.Open();
                    DataTable dt = conn.GetSchema("Databases");
                    conn.Close();
                    return dt;
                }
                catch (Exception ex)
                {
                    if (conn != null)
                        conn.Close();
                    return null;
                }
            }
            public DataTable GetDataTables()
            {
                try
                {
                    string queryTable = "select CAST(0 as bit) as checked,TABLE_NAME, '' as FILE_STRUCT from INFORMATION_SCHEMA.TABLES where TABLE_NAME<>'sysdiagrams' and TABLE_NAME<>'dtproperties' and TABLE_TYPE='BASE TABLE'";// where id = object_id(N'dbo.[tbl_Admin]')";

                    SqlDataAdapter adap = new SqlDataAdapter(queryTable, conn);

                    DataSet dset = new DataSet();
                    adap.Fill(dset);
                    return dset.Tables[0];
                }
                catch
                {
                    return null;
                }
            }
            public void Init(String Server, String Database, String Username, String Pass)
            {
                _connectString = "Data Source=[SERVER];Initial Catalog=[DATABASE]; uid=[USERNAME];pwd=[PASSWORD];Integrated Security=False";
                _connectString = _connectString.Replace("[SERVER]", Server).Replace("[DATABASE]", Database).Replace("[USERNAME]", Username).Replace("[PASSWORD]", Pass);
                conn = new SqlConnection(_connectString);


            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            conn.Init(Server.Text, Database.Text, User.Text, Pass.Text);
            if (conn.TestConnect())
            {
                MessageBox.Show("Connect successful!");
            }
            else MessageBox.Show("Connect fail!");
        }

        private void Database_Enter(object sender, EventArgs e)
        {

        }

        private void Server_Enter(object sender, EventArgs e)
        {

        }

        private void Server_DropDown(object sender, EventArgs e)
        {
            if (Server.DataSource == null)
            {
                Server.DataSource = conn.GetServers();
                Server.DisplayMember = "ServerName";
            }
        }

        private void Database_DropDown(object sender, EventArgs e)
        {
            clsConnectSQL conn = new clsConnectSQL();
            DataTable dt = conn.GetDataBases(Server.Text, User.Text, Pass.Text);
            Database.DataSource = dt;
            Database.DisplayMember = "database_name";
        }


    }
}