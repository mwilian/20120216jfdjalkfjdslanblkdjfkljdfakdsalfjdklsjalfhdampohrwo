using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;

using System.Text;
using System.Windows.Forms;

namespace QueryDesigner
{
    public partial class frmLookup : Form
    {
        string _code = "";
        string _connect = "";

        public string Connect
        {
            get { return _connect; }
            set { _connect = value; }
        }

        public string ReturnCode
        {
            get { return _code; }
            set { _code = value == null ? "" : value.Trim(); }
        }
        public frmLookup()
        {
            InitializeComponent();
        }



        private void btnCancel_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
            Close();
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.OK;
            if (dgvLookup.CurrentRow != null && dgvLookup.CurrentRow.RowIndex >= 0)
            {
                ReturnCode = dgvLookup.CurrentRow.Cells["Code"].Value.ToString();
            }
            Close();
        }

        private void frmLookup_Load(object sender, EventArgs e)
        {
            if (Text == "Table List")
            {
                DataTable kq = new DataTable();
                System.Data.OleDb.OleDbConnection conn = new System.Data.OleDb.OleDbConnection(_connect);
                try
                {

                    conn.Open();
                    kq = conn.GetOleDbSchemaTable(System.Data.OleDb.OleDbSchemaGuid.Tables, new Object[] { null, null, null, "TABLE" });
                    //conn.Close();
                    kq.Columns["TABLE_NAME"].ColumnName = "Code";
                }
                catch { }
                finally { conn.Close(); }
                //dgvLookup.AutoGenerateColumns = false;
                dgvLookup.DataSource = kq;
                
            }
            else if (Text == "View List")
            {
                DataTable kq = new DataTable();
                System.Data.OleDb.OleDbConnection conn = new System.Data.OleDb.OleDbConnection(_connect);
                try
                {

                    conn.Open();
                    kq = conn.GetOleDbSchemaTable(System.Data.OleDb.OleDbSchemaGuid.Tables, new Object[] { null, null, null, "VIEW" });

                    kq.Columns["TABLE_NAME"].ColumnName = "Code";
                }
                catch { }
                finally { conn.Close(); }
                //dgvLookup.AutoGenerateColumns = false;
                dgvLookup.DataSource = kq;
            }
            dgvLookup.AutoSizeColumns();
        }

        private void dgvLookup_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            btnOK_Click(null, null);
        }

        private void dgvLookup_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                btnOK_Click(null, null);
        }
    }
}
