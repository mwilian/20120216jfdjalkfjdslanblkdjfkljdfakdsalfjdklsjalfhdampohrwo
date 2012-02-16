using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace QueryDesigner
{
    public partial class frmValidatedList : Form
    {
        string _sErr = "";
        string _db = "";
        string _usr = "";
        objValidatedList _objReturn;

        public objValidatedList ObjReturn
        {
            get { return _objReturn; }
            set { _objReturn = value; }
        }

        public frmValidatedList(string db, string usr, objValidatedList value)
        {
            InitializeComponent();
            _db = db;
            _usr = usr;
            _objReturn = value;
        }
        public frmValidatedList(string db, string usr)
        {
            InitializeComponent();
            _db = db;
            _usr = usr;
        }
       

        private void frmValidatedList_Load(object sender, EventArgs e)
        {
            BUS.LIST_QDControl qdCtr = new BUS.LIST_QDControl();
            DataTable dt = qdCtr.GetAll_LIST_QD_USER(_db, _usr, ref _sErr);
            ddlQD.DataSource = dt;
            if (_objReturn != null)
            {
                ddlQD.Text = _objReturn.QD;
                ddlFld.Text = _objReturn.Field;
                txtMessage.Text = _objReturn.Message;
            }
        }

        private void multiColumnCombo1_ValueChanged(object sender, EventArgs e)
        {
            BUS.LIST_QDDControl qddCtr = new BUS.LIST_QDDControl();
            DataTable dt = qddCtr.GetALL_LIST_QDD_By_QD_ID(_db, ((DataRowView)ddlQD.SelectedItem).Row["QD_ID"].ToString(), ref _sErr);
            ddlFld.DataSource = dt;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (ddlQD.Text != "" && ddlFld.Text != "")
            {
                DialogResult = DialogResult.OK;
                
                    _objReturn = new objValidatedList(ddlQD.Text, ddlFld.Text, txtMessage.Text);
            }
            else DialogResult = DialogResult.Cancel;
            Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
            Close();
        }


    }
}
