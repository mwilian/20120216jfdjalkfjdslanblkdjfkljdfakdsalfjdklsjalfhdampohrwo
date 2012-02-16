using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace QueryDesigner
{
    public partial class frmRelation : Form
    {
        DataTable _relation = new DataTable("relation");

        public DataTable Relation
        {
            get { return _relation; }
            set { _relation = value; }
        }
        DataTable _original = new DataTable();
        public DataTable DTOriginal
        {
            get { return _original; }
            set { _original = value; }
        }
        DataTable _looup = new DataTable();

        public DataTable DTLooup
        {
            get { return _looup; }
            set { _looup = value; }
        }
        public frmRelation()
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
            Close();
        }

        private void frmRelation_Load(object sender, EventArgs e)
        {
            DataColumn[] col = new DataColumn[] { new DataColumn("Original"), new DataColumn("Lookup") };
            _relation.Columns.AddRange(col);

            dgvRelation.DataSource = _relation;

            foreach (DataRow row in DTOriginal.Rows)
                if (row["type"].ToString() != "S")
                    dgvRelation.RootTable.Columns["Original"].ValueList.Add(row["node"], row["node"].ToString());
            foreach (DataRow row in DTLooup.Rows)
                if (row["type"].ToString() != "S")
                    dgvRelation.RootTable.Columns["Lookup"].ValueList.Add(row["node"], row["node"].ToString());
        }
    }
}
